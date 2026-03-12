import os
import io
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from docxtpl import DocxTemplate
import jinja2
import re
from datetime import datetime
from docx import Document
from docxcompose.composer import Composer

# 引入我们抽离出来的公共工具
from utils.formatters import filter_date, filter_num, parse_dynamic_list


class ResumeGeneratorUI:
    """智能简历引擎 (动态旋转图标 + 极速阻断 + 状态机管控)"""

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

    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

        # 1. 路径选择区
        tk.Label(frame, text="1. Word 简历模板:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0,
                                                                                                sticky=tk.W, pady=5)
        self.entry_tpl = tk.Entry(frame, textvariable=self.tpl_path_var, font=("Microsoft YaHei UI", 10))
        self.entry_tpl.grid(row=0, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="选择模板...", command=self.browse_tpl, width=12).grid(row=0, column=2, pady=5)

        tk.Label(frame, text="2. Excel 人员清单:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0,
                                                                                                 sticky=tk.W, pady=5)
        self.entry_data = tk.Entry(frame, textvariable=self.data_path_var, font=("Microsoft YaHei UI", 10))
        self.entry_data.grid(row=1, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="选择Excel...", command=self.browse_data, width=12).grid(row=1, column=2, pady=5)

        tk.Label(frame, text="3. 简历保存位置:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0,
                                                                                               sticky=tk.W, pady=5)
        self.entry_out = tk.Entry(frame, textvariable=self.out_dir_var, font=("Microsoft YaHei UI", 10))
        self.entry_out.grid(row=2, column=1, padx=10, pady=5, sticky=tk.EW, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.browse_out, width=12).grid(row=2, column=2, pady=5)

        # 4. 控制按钮区
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=20, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1), weight=1)

        self.btn_generate = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_generate.grid(row=0, column=0, padx=10, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=1, padx=10, sticky=tk.EW, ipady=5)

        # 5. 日志区
        tk.Label(frame, text="运行追踪与日志:", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333").grid(row=4,
                                                                                                            column=0,
                                                                                                            sticky=tk.W,
                                                                                                            pady=(10,
                                                                                                                  5))

        self.log_text = scrolledtext.ScrolledText(
            frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50", padx=12, pady=12, relief=tk.FLAT,
            state='disabled'
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

    # ================= 动态旋转动画引擎 =================
    def _animate_btn(self):
        if not self.anim_running:
            return
        frame_char = self.spinner_frames[self.anim_frame % len(self.spinner_frames)]
        self.btn_generate.config(text=f"{frame_char} 引擎高速运转中...")
        self.anim_frame += 1
        self.parent.after(100, self._animate_btn)

    # ================= 按钮状态机管理器 =================
    def update_button_ui(self, state):
        color_start_active = "#4CAF50"  # 墨绿色
        color_stop_active = "#DC3545"  # 警告红
        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray = "#E0E0E0"

        if state == "idle":
            self.anim_running = False  # 确保动画关闭
            self.btn_generate.config(bg=color_start_active, fg="white", text="▶ 开始智能生成", state=tk.NORMAL,
                                     command=self.start_generation_thread)
            self.btn_cancel.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏹ 停止生成", state=tk.DISABLED)

        elif state == "running":
            # 【体验提升】：保持绿色高亮，并触发旋转动画
            self.btn_generate.config(bg=color_start_active, fg="white", state=tk.NORMAL, command=lambda: None)
            self.btn_cancel.config(bg=color_enabled_gray, fg="#333", text="⏹ 停止生成", state=tk.NORMAL,
                                   command=self.cancel_processing)

            # 启动动画环
            if not self.anim_running:
                self.anim_running = True
                self.anim_frame = 0
                self._animate_btn()

        elif state == "stopping":
            self.anim_running = False  # 瞬间停止动画
            self.btn_generate.config(bg=color_disabled_bg, fg=color_disabled_fg, text="🛑 正在踩刹车...",
                                     state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="⏹ 正在清理现场...", state=tk.NORMAL,
                                   command=lambda: None)

    # ================= UI 交互流 =================
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
            if auto_scroll:
                self.log_text.see(tk.END)
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
        if messagebox.askyesno("确认停止", "确定要立刻中断简历生成吗？\n中断后将不会拼接待汇总的文档。"):
            self._stop_event.set()
            self.update_button_ui("stopping")
            self.log_to_ui("🛑 收到阻断指令，正在丢弃已生成的数据并清理缓存...", "WARNING", auto_scroll=False)

    def is_file_locked(self, filepath):
        """检测文件是否被其他程序（如 Word/Excel）打开占用"""
        if not os.path.exists(filepath):
            return False
        try:
            # 尝试重命名文件自身，如果被 Office 打开，Windows 会拒绝并抛出 PermissionError
            os.rename(filepath, filepath)
            return False
        except OSError:
            return True

    def start_generation_thread(self):
        tpl = self.tpl_path_var.get().strip().strip('"').strip("'")
        data = self.data_path_var.get().strip().strip('"').strip("'")
        out = self.out_dir_var.get().strip().strip('"').strip("'")

        if not tpl or not data or not out:
            messagebox.showwarning("提示", "请先填入完整的模板路径、Excel 人员清单和输出目录！")
            return

        if not os.path.exists(tpl) or not os.path.exists(data):
            messagebox.showerror("错误", "模板或Excel 人员清单文件不存在，请检查路径是否正确。")
            return

        # ================= [新增] 共享冲突前置拦截 =================
        locked_files = []
        if self.is_file_locked(tpl):
            locked_files.append("Word 简历模板")
        if self.is_file_locked(data):
            locked_files.append("Excel 人员清单")

        if locked_files:
            msg = f"检测到以下文件正在被其他程序（如 Word/Excel）打开：\n\n" \
                  f"【 {'、'.join(locked_files)} 】\n\n" \
                  f"请先关闭这些文件后再点击开始，以免产生共享冲突！"
            messagebox.showwarning("文件被占用", msg)
            return
        # ==========================================================

        self._stop_event.clear()
        self.update_button_ui("running")

        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        threading.Thread(target=self.core_generate_logic, args=(tpl, data, out), daemon=True).start()

    # ================= 核心业务引擎 =================
    def core_generate_logic(self, tpl_path, data_path, out_dir):
        self.log_to_ui("=== [阶段1] 开始解析 Excel 人员清单 ===", "HIGHLIGHT")
        try:
            # ================= [修改点 1: 内存加载 Excel，防占用] =================
            with open(data_path, 'rb') as f:
                excel_bytes = f.read()  # 瞬间将整个 Excel 读入内存

            # 使用内存数据加载，彻底切断与硬盘文件的联系
            with pd.ExcelFile(io.BytesIO(excel_bytes)) as xls:
                sheet_names = xls.sheet_names

            # 1. 尝试精准狙击
            target_sheet = '实施人员清单'
            if target_sheet in sheet_names:
                df = pd.read_excel(data_path, sheet_name=target_sheet, dtype=str).fillna('')
                self.log_to_ui(f"  └─ 已精准定位到 '{target_sheet}' 表。")
                sheet_to_read = target_sheet
            else:
                # 2. 找不到指定名称，触发柔性降级，直接强读第 1 个 Sheet
                df = pd.read_excel(data_path, sheet_name=0, dtype=str).fillna('')
                self.log_to_ui(f"  └─ 未找到 '{target_sheet}'，默认读取第一个表: '{sheet_names[0]}'。")
                sheet_to_read = 0

            # 【防呆】原始表头重复拦截机制
            raw_header_row = pd.read_excel(data_path, sheet_name=sheet_to_read, header=None, nrows=1).values[0]
            raw_headers = [str(x).strip() for x in raw_header_row if pd.notna(x) and str(x).strip() != '']

            seen, duplicates = set(), set()
            for h in raw_headers:
                if h in seen: duplicates.add(h)
                seen.add(h)

            if duplicates:
                err_msg = (
                    f"Excel 中存在重复的表头名称：\n【 {', '.join(duplicates)} 】\n\n"
                    f"💡 存在重复表头会导致数据混淆。请在 Excel 中将它们重命名为唯一名称。"
                )
                self.log_to_ui("❌ [数据拦截] 发现重复表头，已终止运行。", "ERROR")
                messagebox.showwarning("表头查重失败", err_msg)
                self.parent.after(0, lambda: self.update_button_ui("idle"))
                return

            self.log_to_ui(f"  └─ 成功读取数据，共发现 {len(df)} 条人员记录。", "SUCCESS")

        except Exception as e:
            self.log_to_ui(f"❌ 读取 Excel 失败: {str(e)}", "ERROR")
            self.parent.after(0, lambda: self.update_button_ui("idle"))
            return

        # 初始化 Jinja 渲染引擎
        jinja_env = jinja2.Environment()
        jinja_env.filters['date'] = filter_date
        jinja_env.filters['num'] = filter_num

        success_count = 0
        error_list = []
        generated_files = []

        self.log_to_ui("\n=== [阶段2] 极速渲染单人简历模板 ===", "HIGHLIGHT")

        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        # ================= [修改点 2: 内存加载 Word 模板，防占用并提速] =================
        try:
            with open(tpl_path, 'rb') as f:
                tpl_bytes = f.read()  # 把模板一次性读入内存
        except Exception as e:
            self.log_to_ui(f"❌ 读取 Word 模板失败: {str(e)}", "ERROR")
            self.parent.after(0, lambda: self.update_button_ui("idle"))
            return
        # ==============================================================================

        for index, row in df.iterrows():
            if self._stop_event.is_set():
                break

            name = str(row.get('姓名', '')).strip()
            if not name: name = f"未命名_第{index + 2}行"

            try:
                context = {}
                for col_name in df.columns:
                    col_str = str(col_name).strip()
                    cell_val = str(row[col_name]).strip()
                    # 1. 原始单元格字符串
                    context[col_str] = cell_val

                    # 2. 恢复原版的高级解析：保留字典结构
                    # 专门用于截图中的复杂表格渲染 ({% tr %})，支持取 {{ item.c1 }} 等
                    context[f"{col_str}_list"] = parse_dynamic_list(cell_val)

                    # 3. 新增【纯净行】解析：剔除多余符号的纯文本列表
                    # 专门用于上一个问题中的“动态符号+换行”排版 ({% p %})
                    clean_lines = [line.strip() for line in cell_val.split('\n') if line.strip()]
                    # 剔除行首可能自带的序号或黑点

                    final_lines = [re.sub(r'^[0-9]+[、\.．\s]+|^[●■◆✔-]\s*', '', line) for line in clean_lines]
                    context[f"{col_str}_lines"] = final_lines

                # ================= [修改点 3: 从内存流中渲染] =================
                # 这样做不再需要每次都去读硬盘，彻底消灭文件占用！
                doc = DocxTemplate(io.BytesIO(tpl_bytes))
                # ============================================================

                doc.render(context, jinja_env=jinja_env)

                time_str = datetime.now().strftime("%H%M%S")
                safe_name = "".join(c for c in name if c not in r'\/:*?"<>|')
                save_path = os.path.join(out_dir, f"~temp_{safe_name}_{time_str}.docx")

                doc.save(save_path)
                generated_files.append(save_path)
                success_count += 1

                self.log_to_ui(f"  └─ 成功生成: {name} ({success_count}/{len(df)})")

            except jinja2.exceptions.TemplateSyntaxError as je:
                err_msg = (
                    f"❌ [致命模板错误] 渲染 '{name}' 时崩溃！\n"
                    f"  └─ 原因: Word模板中使用了非法的标签 -> {str(je)}\n"
                    f"  └─ 建议: 请检查 {{{{ }}}} 标签，不要将带括号或特殊字符的表头直接放入。\n"
                    f"⚠️ 已自动触发熔断，终止后续所有人员的生成。"
                )
                self.log_to_ui(err_msg, "ERROR")
                error_list.append(name)
                break

            except Exception as e:
                self.log_to_ui(f"❌ [错误] {name} 生成失败: {str(e)}", "ERROR")
                error_list.append(name)

        # ================= 中断清场 or 无缝合并 =================
        if self._stop_event.is_set():
            self.log_to_ui("\n=== [清场] 清理未完成的临时碎片 ===", "WARNING", auto_scroll=False)
            for f in generated_files:
                try:
                    os.remove(f)
                except:
                    pass
            self.log_to_ui("🧹 清理完毕，引擎已安全退出。", "SUCCESS", auto_scroll=False)
            self.parent.after(0, lambda: self.update_button_ui("idle"))
            return

        if len(generated_files) > 0:
            self.log_to_ui("\n=== [阶段3] 拼接总文档与清理缓存 ===", "HIGHLIGHT")
            try:
                self.log_to_ui("  └─ 正在将所有人员无缝装订成册，请稍候...")
                master = Document(generated_files[0])
                composer = Composer(master)

                for file_path in generated_files[1:]:
                    doc_append = Document(file_path)
                    composer.append(doc_append)

                merged_name = f"智能汇总_共{len(generated_files)}人简历_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                merged_path = os.path.join(out_dir, merged_name)

                if os.path.exists(merged_path):
                    os.remove(merged_path)

                composer.save(merged_path)

                self.log_success_link("★ 拼接大满贯！点击立即打开验收 ➔ ", merged_name, merged_path)

                clean_count = 0
                for file_path in generated_files:
                    try:
                        os.remove(file_path)
                        clean_count += 1
                    except:
                        pass
                self.log_to_ui(f"  └─ 痕迹抹除完毕，共清理 {clean_count} 个中间文件。", "INFO")

            except PermissionError:
                self.log_to_ui("❌ [合并失败] 您可能正在用 Word 打开同名文件，无法覆盖写入！", "ERROR")
                self.log_to_ui("💡 但为了数据安全，所有单人的临时简历已为您保留在输出目录中。", "WARNING")
            except Exception as e:
                self.log_to_ui(f"❌ [合并失败] 发生异常: {str(e)}", "ERROR")
                self.log_to_ui("💡 单人简历缓存已保留供排查。", "WARNING")

        self.log_to_ui("\n=== 🏁 任务结束 ===", "HIGHLIGHT")
        self.log_to_ui(f"战报: 成功 {success_count} 人 | 失败 {len(error_list)} 人", "SUCCESS")

        self.parent.after(0, lambda: self.update_button_ui("idle"))