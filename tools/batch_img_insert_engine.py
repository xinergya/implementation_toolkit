import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import openpyxl
from datetime import datetime
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
from docx.shared import Cm

from utils.ui_base import BaseEngineUI


# ==========================================
# 1. 核心处理引擎类 (Engine) - 支持彩色等级日志
# ==========================================
class BatchImageInsertEngine:
    def __init__(self, excel_path, template_path, shebao_dir, gongzi_dir, output_dir, log_callback=None):
        self.excel_path = excel_path
        self.template_path = template_path
        self.shebao_dir = shebao_dir
        self.gongzi_dir = gongzi_dir
        self.output_dir = output_dir
        self.log = log_callback if log_callback else (lambda msg, level="INFO", **kwargs: print(f"[{level}] {msg}"))

    def find_images_fuzzy(self, name, search_dir):
        valid_exts = ('.jpg', '.jpeg', '.png')
        matched_imgs = []
        if not search_dir or not os.path.exists(search_dir):
            return matched_imgs

        for root, dirs, files in os.walk(search_dir):
            for file in files:
                file_lower = file.lower()
                if name in file and file_lower.endswith(valid_exts):
                    matched_imgs.append(os.path.join(root, file))
        matched_imgs.sort()
        return matched_imgs

    def execute(self):
        try:
            self.log("=== [阶段1] 开始智能解析 Excel 数据 ===", "HIGHLIGHT")
            df = pd.read_excel(self.excel_path, dtype=str).fillna('')
            df.columns = [str(c).strip() for c in df.columns]

            name_col = next((col for col in df.columns if '姓名' in col), None)
            xuhao_col = next((col for col in df.columns if '序号' in col), None)

            if not name_col:
                self.log("❌ [错误] 智能嗅探失败，Excel表头中未找到包含“姓名”字样的列！", "ERROR")
                return False

            self.log(f"  └─ 智能匹配核心字段: 姓名列 -> [{name_col}], 序号列 -> [{xuhao_col or '未找到,采用默认序号'}]", "INFO")

            os.makedirs(self.output_dir, exist_ok=True)
            self.log(f"  └─ 输出目录就绪: {self.output_dir}", "SUCCESS")

            total = len(df)
            success_count = 0
            generated_files = []

            self.log("\n=== [阶段2] 批量插入图片并生成独立文档 ===", "HIGHLIGHT")
            for index, row in df.iterrows():
                context = {}
                for col_name in df.columns:
                    context[col_name] = str(row[col_name]).strip()

                name = context.get(name_col, '')
                xuhao = context.get(xuhao_col, str(index + 1)) if xuhao_col else str(index + 1)

                if not name:
                    continue

                self.log(f"[{index + 1}/{total}] 正在生成: {name}", "INFO")
                tpl = DocxTemplate(self.template_path)

                shebao_imgs = self.find_images_fuzzy(name, self.shebao_dir)
                if shebao_imgs:
                    sub_sb = tpl.new_subdoc()
                    for img_path in shebao_imgs:
                        sub_sb.add_paragraph().add_run().add_picture(img_path, width=Cm(15))
                    context['社保图片'] = sub_sb
                    self.log(f"  └─ [社保] 成功匹配到 {len(shebao_imgs)} 张图片", "SUCCESS")
                else:
                    context['社保图片'] = "【未找到社保图片】"
                    self.log(f"  └─ [社保] 未找到对应图片", "WARNING")

                gongzi_imgs = self.find_images_fuzzy(name, self.gongzi_dir)
                if gongzi_imgs:
                    sub_gz = tpl.new_subdoc()
                    for img_path in gongzi_imgs:
                        sub_gz.add_paragraph().add_run().add_picture(img_path, width=Cm(15))
                    context['工资图片'] = sub_gz
                    self.log(f"  └─ [工资] 成功匹配到 {len(gongzi_imgs)} 张图片", "SUCCESS")
                else:
                    context['工资图片'] = "【未找到工资图片】"
                    self.log(f"  └─ [工资] 未找到对应图片", "WARNING")

                tpl.render(context)
                out_file = os.path.join(self.output_dir, f"{name}_证明.docx")
                tpl.save(out_file)
                generated_files.append(out_file)
                success_count += 1

            self.log("\n=== [阶段3] 汇总文档并回写 Excel 索引 ===", "HIGHLIGHT")
            merged_out = os.path.join(self.output_dir, "00_合并汇总_人员证明.docx")
            if success_count > 0:
                self.log(f"  └─ 开始无缝合并 {success_count} 份独立文档...", "INFO")
                try:
                    master = Document(generated_files[0])
                    composer = Composer(master)
                    for file_path in generated_files[1:]:
                        master.add_page_break()
                        composer.append(Document(file_path))
                    composer.save(merged_out)
                    self.log(f"  └─ 汇总文档生成成功！", "SUCCESS")
                except Exception as e:
                    self.log(f"  └─ [报错] 合并过程发生异常: {str(e)}", "ERROR")

            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active

                name_col_idx = None
                xuhao_col_idx = None

                for cell in ws[1]:
                    val = str(cell.value).strip() if cell.value else ""
                    if val == name_col:
                        name_col_idx = cell.column
                    elif xuhao_col and val == xuhao_col:
                        xuhao_col_idx = cell.column

                if not name_col_idx:
                    self.log(f"  └─ [报错] 未能定位姓名列，跳过更新表结构", "ERROR")
                else:
                    for r in range(2, ws.max_row + 1):
                        cell_name = ws.cell(row=r, column=name_col_idx).value
                        cell_xuhao = ws.cell(row=r, column=xuhao_col_idx).value if xuhao_col_idx else r - 1

                        if not cell_name or not str(cell_name).strip():
                            continue

                        str_shebao = f"{cell_xuhao}.1 {cell_name}-参保证明"
                        str_gongzi = f"{cell_xuhao}.2 {cell_name}-工资证明"

                        ws.cell(row=r, column=6).value = str_shebao
                        ws.cell(row=r, column=8).value = str_gongzi

                    updated_excel_path = os.path.join(self.output_dir, "00_人员清单_已自动更新目录.xlsx")
                    wb.save(updated_excel_path)
                    self.log(f"  └─ 目录索引已精准填充至表格 F 列与 H 列！", "SUCCESS")
            except Exception as e:
                self.log(f"  └─ [警告] 更新 Excel 发生异常 (文件可能被占用): {str(e)}", "WARNING")

            self.log(f"\n=== 🏁 任务结束 (共处理 {success_count} 人) ===", "HIGHLIGHT")
            return merged_out if success_count > 0 else None

        except Exception as e:
            self.log(f"❌ 核心引擎发生严重异常: {str(e)}", "ERROR")
            return None


# ==========================================
# 2. 界面交互类 (UI)
# ==========================================
class BatchImageInsertUI(BaseEngineUI):
    def __init__(self, parent_frame):
        super().__init__(parent_frame)
        self.latest_merged_path = ""
        self.setup_ui()

    def setup_ui(self):
        self.parent.columnconfigure(1, weight=1)
        self.parent.rowconfigure(6, weight=1)

        # --- 1-4. 路径选择区 ---
        ttk.Label(self.parent, text="人员清单(Excel):").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.excel_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.excel_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_excel).grid(row=0, column=2, padx=10, pady=10)

        ttk.Label(self.parent, text="Word模板(.docx):").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.tpl_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.tpl_entry.grid(row=1, column=1, padx=5, pady=10, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_tpl).grid(row=1, column=2, padx=10, pady=10)

        ttk.Label(self.parent, text="社保图片目录:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.shebao_dir_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.shebao_dir_entry.grid(row=2, column=1, padx=5, pady=10, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_shebao_dir).grid(row=2, column=2, padx=10, pady=10)

        ttk.Label(self.parent, text="工资图片目录:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.gongzi_dir_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.gongzi_dir_entry.grid(row=3, column=1, padx=5, pady=10, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_gongzi_dir).grid(row=3, column=2, padx=10, pady=10)

        # --- 5. 执行按钮 ---
        self.execute_btn = tk.Button(self.parent, text="▶ 全自动合图并填表", bg="#4CAF50", fg="white",
                                     font=("Microsoft YaHei UI", 11, "bold"), command=self.on_execute)
        self.execute_btn.grid(row=4, column=0, columnspan=3, pady=15, sticky=tk.EW, padx=20, ipady=5)

        # --- 6. 运行追踪与日志区 (完全对齐工业级标准) ---
        tk.Label(self.parent, text="运行追踪与日志:", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333").grid(row=5, column=0, sticky=tk.W, padx=15, pady=(10, 5))

        self.log_text = scrolledtext.ScrolledText(
            self.parent, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
            padx=12, pady=12, relief=tk.FLAT, state='disabled'
        )
        self.log_text.grid(row=6, column=0, columnspan=3, padx=15, pady=5, sticky=tk.NSEW)

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

        self.log_to_ui("✨ 智能合图填表引擎已就绪！", "HIGHLIGHT")

    # --- 对齐的日志与交互方法 ---
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

    # --- 文件选择与调度器 ---
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx;*.xls")])
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)

    def select_tpl(self):
        path = filedialog.askopenfilename(filetypes=[("Word 模板", "*.docx")])
        if path:
            self.tpl_entry.delete(0, tk.END)
            self.tpl_entry.insert(0, path)

    def select_shebao_dir(self):
        path = filedialog.askdirectory(title="选择社保图片目录")
        if path:
            self.shebao_dir_entry.delete(0, tk.END)
            self.shebao_dir_entry.insert(0, path)

    def select_gongzi_dir(self):
        path = filedialog.askdirectory(title="选择工资图片目录")
        if path:
            self.gongzi_dir_entry.delete(0, tk.END)
            self.gongzi_dir_entry.insert(0, path)

    def on_execute(self):
        excel_path = self.excel_entry.get().strip()
        tpl_path = self.tpl_entry.get().strip()
        shebao_dir = self.shebao_dir_entry.get().strip()
        gongzi_dir = self.gongzi_dir_entry.get().strip()

        if not all([excel_path, tpl_path, shebao_dir, gongzi_dir]):
            messagebox.showwarning("参数缺失", "请先填写所有必须的文件和目录路径！")
            return

        output_dir = os.path.join(os.path.dirname(tpl_path), "已生成_人员证明")

        self.execute_btn.config(state="disabled", text="引擎高速运转中...", bg="#E9ECEF", fg="#ADB5BD")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        self.log_to_ui("🚀 启动智能合图填表流水线...", "HIGHLIGHT")

        engine = BatchImageInsertEngine(
            excel_path=excel_path,
            template_path=tpl_path,
            shebao_dir=shebao_dir,
            gongzi_dir=gongzi_dir,
            output_dir=output_dir,
            log_callback=self.log_to_ui
        )

        def run_thread():
            merged_file_path = engine.execute()

            # 恢复按钮状态
            def restore_btn():
                self.execute_btn.config(state="normal", text="▶ 全自动合图并填表", bg="#4CAF50", fg="white")

            self.parent.after(0, restore_btn)

            # 若成功生成最终文档，触发超链接日志
            if merged_file_path:
                file_name = os.path.basename(merged_file_path)
                self.log_success_link("★ 任务全线竣工！点击立即打开验收 ➔ ", file_name, merged_file_path)

        threading.Thread(target=run_thread, daemon=True).start()