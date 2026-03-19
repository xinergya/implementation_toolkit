import os
import io
import re
import threading
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime

# Word 相关依赖
from docxtpl import DocxTemplate
import jinja2
from docx import Document
from docxcompose.composer import Composer

# 引入抽离出来的公共工具 (请确保你的 utils 目录结构正确)
from utils.formatters import filter_date, filter_num, parse_dynamic_list


class ResumeGeneratorUI:
    """智能简历引擎 (V2.0 工业级重构版：全内存I/O + 系统级临时沙箱 + OOD拆分)"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        self._stop_event = threading.Event()

        self.tpl_path_var = tk.StringVar()
        self.data_path_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()

        # UI 动画控制变量
        self.anim_running = False
        self.anim_frame = 0
        self.spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

        self.setup_ui()

    # ================= 1. UI 构建与布局 =================
    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

        # 路径选择区
        tk.Label(frame, text="1. Word 简历模板:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entry_tpl = tk.Entry(frame, textvariable=self.tpl_path_var, font=("Microsoft YaHei UI", 10))
        self.entry_tpl.grid(row=0, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="选择模板...", command=self.browse_tpl, width=12).grid(row=0, column=2, pady=5)

        tk.Label(frame, text="2. Excel 人员清单:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.entry_data = tk.Entry(frame, textvariable=self.data_path_var, font=("Microsoft YaHei UI", 10))
        self.entry_data.grid(row=1, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="选择Excel...", command=self.browse_data, width=12).grid(row=1, column=2, pady=5)

        tk.Label(frame, text="3. 简历保存位置:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.entry_out = tk.Entry(frame, textvariable=self.out_dir_var, font=("Microsoft YaHei UI", 10))
        self.entry_out.grid(row=2, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.browse_out, width=12).grid(row=2, column=2, pady=5)

        # 控制按钮区
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=20, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1), weight=1)

        self.btn_generate = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_generate.grid(row=0, column=0, padx=10, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=1, padx=10, sticky=tk.EW, ipady=5)

        # 日志区
        tk.Label(frame, text="运行追踪与日志:", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333").grid(row=4, column=0, sticky=tk.W, pady=(10, 5))

        self.log_text = scrolledtext.ScrolledText(
            frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50", padx=12, pady=12, relief=tk.FLAT, state='disabled'
        )
        self.log_text.grid(row=5, column=0, columnspan=3, pady=5, sticky=tk.NSEW)

        # 富文本标签样式
        self.log_text.tag_config("INFO", foreground="#2C3E50")
        self.log_text.tag_config("SUCCESS", foreground="#198754", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("WARNING", foreground="#D35400", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("ERROR", foreground="#DC3545", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("HIGHLIGHT", foreground="#8E44AD", font=("Microsoft YaHei UI", 11, "bold"))
        self.log_text.tag_config("LINK", foreground="#0078D7", underline=True, font=("Microsoft YaHei UI", 11, "bold"))

        self.log_text.tag_bind("LINK", "<Button-1>", self.open_generated_file)
        self.log_text.tag_bind("LINK", "<Enter>", lambda e: self.log_text.config(cursor="hand2"))
        self.log_text.tag_bind("LINK", "<Leave>", lambda e: self.log_text.config(cursor=""))

        self.log_to_ui("✨ 简历批量生成与无缝合并引擎已就绪！", "HIGHLIGHT")
        self.update_button_ui("idle")

    # ================= 2. 交互流与状态机 =================
    def _animate_btn(self):
        if not self.anim_running:
            return
        frame_char = self.spinner_frames[self.anim_frame % len(self.spinner_frames)]
        self.btn_generate.config(text=f"{frame_char} 引擎高速运转中...")
        self.anim_frame += 1
        self.parent.after(100, self._animate_btn)

    def update_button_ui(self, state):
        color_start_active = "#4CAF50"
        color_stop_active = "#DC3545"
        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray = "#E0E0E0"

        if state == "idle":
            self.anim_running = False
            self.btn_generate.config(bg=color_start_active, fg="white", text="▶ 开始智能生成", state=tk.NORMAL, command=self.start_generation_thread)
            self.btn_cancel.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏹ 停止生成", state=tk.DISABLED)
        elif state == "running":
            self.btn_generate.config(bg=color_start_active, fg="white", state=tk.NORMAL, command=lambda: None)
            self.btn_cancel.config(bg=color_enabled_gray, fg="#333", text="⏹ 停止生成", state=tk.NORMAL, command=self.cancel_processing)
            if not self.anim_running:
                self.anim_running = True
                self.anim_frame = 0
                self._animate_btn()
        elif state == "stopping":
            self.anim_running = False
            self.btn_generate.config(bg=color_disabled_bg, fg=color_disabled_fg, text="🛑 正在踩刹车...", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="⏹ 正在清理现场...", state=tk.NORMAL, command=lambda: None)

    def browse_tpl(self):
        path = filedialog.askopenfilename(title="选择 Word 模板", filetypes=[("Word 文档", "*.docx")])
        if path: self.tpl_path_var.set(path)

    def browse_data(self):
        path = filedialog.askopenfilename(title="选择 Excel 人员清单", filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if path:
            self.data_path_var.set(path)
            self.out_dir_var.set(os.path.dirname(path))

    def browse_out(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path: self.out_dir_var.set(path)

    def open_generated_file(self, event):
        if hasattr(self, 'latest_merged_path') and os.path.exists(self.latest_merged_path):
            try:
                os.startfile(self.latest_merged_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件，可能文件已被移动或删除。\n{e}")

    def log_to_ui(self, message, level="INFO", auto_scroll=True):
        def update():
            self.log_text.config(state='normal')
            timestamp = datetime.now().strftime("[%H:%M:%S] ")
            self.log_text.insert(tk.END, f"{timestamp}{message}\n", level)
            if auto_scroll: self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
            self.parent.update_idletasks()

        self.parent.after(0, update)

    def log_success_link(self, highlight_text, file_name, file_path):
        self.latest_merged_path = file_path

        def update():
            self.log_text.config(state='normal')
            timestamp = datetime.now().strftime("\n[%H:%M:%S] ")
            self.log_text.insert(tk.END, timestamp)
            self.log_text.insert(tk.END, highlight_text, "HIGHLIGHT")
            self.log_text.insert(tk.END, f" {file_name}\n", "LINK")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
            self.parent.update_idletasks()

        self.parent.after(0, update)

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", "确定要立刻中断简历生成吗？\n中断后将丢弃未完成的碎片。"):
            self._stop_event.set()
            self.update_button_ui("stopping")
            self.log_to_ui("🛑 收到阻断指令，正在安全停止任务...", "WARNING", auto_scroll=False)

    def is_file_locked(self, filepath):
        if not os.path.exists(filepath): return False
        try:
            os.rename(filepath, filepath)
            return False
        except OSError:
            return True

    # ================= 3. 核心业务调度中心 (Orchestrator) =================
    def start_generation_thread(self):
        tpl = self.tpl_path_var.get().strip().strip('"').strip("'")
        data = self.data_path_var.get().strip().strip('"').strip("'")
        out = self.out_dir_var.get().strip().strip('"').strip("'")

        if not tpl or not data or not out:
            messagebox.showwarning("提示", "请先填入完整的模板路径、Excel清单和输出目录！")
            return
        if not os.path.exists(tpl) or not os.path.exists(data):
            messagebox.showerror("错误", "模板或Excel清单文件不存在，请检查路径是否正确。")
            return

        locked_files = []
        if self.is_file_locked(tpl): locked_files.append("Word 简历模板")
        if self.is_file_locked(data): locked_files.append("Excel 人员清单")
        if locked_files:
            messagebox.showwarning("文件被占用", f"检测到以下文件被其他程序占用：\n【 {'、'.join(locked_files)} 】\n请先关闭后再试！")
            return

        self._stop_event.clear()
        self.update_button_ui("running")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        threading.Thread(target=self.core_generate_logic, args=(tpl, data, out), daemon=True).start()

    def core_generate_logic(self, tpl_path, data_path, out_dir):
        """【架构重构】：SRP 单一职责调度器，彻底告别上帝函数"""
        try:
            # 阶段 1：真·全内存加载 Excel 数据
            df = self._load_excel_data(data_path)
            if df is None: return

            # 阶段 2：使用系统级 TemporaryDirectory 作为防爆隔离沙箱
            with tempfile.TemporaryDirectory() as temp_dir:
                generated_files = self._render_individual_resumes(df, tpl_path, temp_dir)

                if self._stop_event.is_set():
                    self.log_to_ui("🧹 已成功阻断，临时沙箱由系统自动销毁。", "SUCCESS")
                    return

                if not generated_files:
                    self.log_to_ui("⚠️ 未生成任何有效的简历文件。", "WARNING")
                    return

                # 阶段 3：无缝合并输出
                self._merge_documents(generated_files, out_dir)

            self.log_to_ui("\n=== 🏁 任务结束 ===", "HIGHLIGHT")

        except Exception as e:
            self.log_to_ui(f"❌ 核心引擎发生严重异常: {str(e)}", "ERROR")
        finally:
            self.parent.after(0, lambda: self.update_button_ui("idle"))

    # ================= 4. 子模块：纯净数据读取层 =================
    def _load_excel_data(self, data_path):
        self.log_to_ui("=== [阶段1] 内存极速解析 Excel ===", "HIGHLIGHT")
        try:
            # 彻底切断与硬盘的联系
            with open(data_path, 'rb') as f:
                excel_bytes = f.read()
            excel_io = io.BytesIO(excel_bytes)

            xls_obj = pd.ExcelFile(excel_io)
            sheet_names = xls_obj.sheet_names

            target_sheet = '实施人员清单'
            if target_sheet in sheet_names:
                df = pd.read_excel(xls_obj, sheet_name=target_sheet, dtype=str).fillna('')
                sheet_to_read = target_sheet
                self.log_to_ui(f"  └─ 已精准定位到 '{target_sheet}' 表。")
            else:
                df = pd.read_excel(xls_obj, sheet_name=0, dtype=str).fillna('')
                sheet_to_read = 0
                self.log_to_ui(f"  └─ 默认读取首个表: '{sheet_names[0]}'。")

            # 表头查重（复用已加载的 xls_obj，拒绝二次读盘）
            raw_header_row = pd.read_excel(xls_obj, sheet_name=sheet_to_read, header=None, nrows=1).values[0]
            raw_headers = [str(x).strip() for x in raw_header_row if pd.notna(x) and str(x).strip() != '']

            seen, duplicates = set(), set()
            for h in raw_headers:
                if h in seen: duplicates.add(h)
                seen.add(h)

            if duplicates:
                self.log_to_ui("❌ [数据拦截] 发现重复表头，已终止运行。", "ERROR")
                messagebox.showwarning("表头查重失败", f"Excel中存在重复表头：\n【{', '.join(duplicates)}】\n请修正！")
                return None

            self.log_to_ui(f"  └─ 成功装载 {len(df)} 条人员记录。", "SUCCESS")
            return df
        except Exception as e:
            self.log_to_ui(f"❌ 读取 Excel 失败: {str(e)}", "ERROR")
            return None

    # ================= 5. 子模块：高性能渲染引擎 =================
    def _render_individual_resumes(self, df, tpl_path, temp_dir):
        self.log_to_ui("\n=== [阶段2] 渲染单人模板 (临时沙箱) ===", "HIGHLIGHT")
        generated_files = []
        error_list = []

        try:
            with open(tpl_path, 'rb') as f:
                tpl_bytes = f.read()
        except Exception as e:
            self.log_to_ui(f"❌ 读取 Word 模板失败: {str(e)}", "ERROR")
            return []

        jinja_env = jinja2.Environment()
        jinja_env.filters['date'] = filter_date
        jinja_env.filters['num'] = filter_num

        # 【性能优化】：将高频使用的正则表达式提纯至循环外预编译
        clean_pattern = re.compile(r'^[0-9]+[、\.．\s]+|^[●■◆✔-]\s*')

        success_count = 0
        for index, row in df.iterrows():
            if self._stop_event.is_set(): break

            name = str(row.get('姓名', '')).strip()
            if not name: name = f"未命名_第{index + 2}行"

            try:
                context = {}
                for col_name in df.columns:
                    col_str = str(col_name).strip()
                    cell_val = str(row[col_name]).strip()

                    context[col_str] = cell_val
                    context[f"{col_str}_list"] = parse_dynamic_list(cell_val)

                    clean_lines = [line.strip() for line in cell_val.split('\n') if line.strip()]
                    # 使用预编译的正则，性能成倍提升
                    context[f"{col_str}_lines"] = [clean_pattern.sub('', line) for line in clean_lines]

                # 纯内存 IO 装载模板，防锁定防崩溃
                doc = DocxTemplate(io.BytesIO(tpl_bytes))
                doc.render(context, jinja_env=jinja_env)

                safe_name = "".join(c for c in name if c not in r'\/:*?"<>|')
                save_path = os.path.join(temp_dir, f"{safe_name}_{index}.docx")
                doc.save(save_path)

                generated_files.append(save_path)
                success_count += 1
                self.log_to_ui(f"  └─ 成功生成: {name} ({success_count}/{len(df)})")

            except jinja2.exceptions.TemplateSyntaxError as je:
                self.log_to_ui(f"❌ [致命模板错误] 渲染 '{name}' 崩溃 -> {str(je)}", "ERROR")
                error_list.append(name)
                break
            except Exception as e:
                self.log_to_ui(f"❌ [错误] {name} 生成失败: {str(e)}", "ERROR")
                error_list.append(name)

        self.log_to_ui(f"\n沙箱渲染进度: 成功 {success_count} 人 | 失败 {len(error_list)} 人", "INFO")
        return generated_files

    # ================= 6. 子模块：文档无缝合并层 =================
    def _merge_documents(self, file_list, out_dir):
        self.log_to_ui("\n=== [阶段3] 拼接总文档并输出 ===", "HIGHLIGHT")
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        try:
            self.log_to_ui("  └─ 正在将所有人员无缝装订成册，请稍候...")
            master = Document(file_list[0])
            composer = Composer(master)

            for file_path in file_list[1:]:
                doc_append = Document(file_path)
                composer.append(doc_append)

            merged_name = f"智能汇总_共{len(file_list)}人简历_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            merged_path = os.path.join(out_dir, merged_name)

            # 防护：强行删除同名旧文件
            if os.path.exists(merged_path):
                os.remove(merged_path)

            composer.save(merged_path)
            self.log_success_link("★ 拼接大满贯！点击立即打开验收 ➔ ", merged_name, merged_path)

        except PermissionError:
            self.log_to_ui("❌ [合并失败] 您正在用 Word 打开同名汇总文件，无法覆盖写入！", "ERROR")
        except Exception as e:
            self.log_to_ui(f"❌ [合并失败] 发生异常: {str(e)}", "ERROR")