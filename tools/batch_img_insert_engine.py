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
from utils.image_types import (
    load_image_types, save_image_types, parse_keywords, match_keywords,
    extract_template_variables, find_person_images, match_types_for_person,
    prepare_image, get_excel_columns, identify_image_type_variables,
)


# ==========================================
# 1. 核心处理引擎类 (Engine)
# ==========================================
class BatchImageInsertEngine:
    def __init__(self, excel_path, template_path, image_types, output_dir, log_callback=None):
        self.excel_path = excel_path
        self.template_path = template_path
        self.image_types = [t for t in image_types if t.get("name")]
        self.output_dir = output_dir
        self.log = log_callback if log_callback else (lambda msg, level="INFO", **kwargs: print(f"[{level}] {msg}"))

    def execute(self):
        try:
            self.log("=== [阶段1] 开始解析 Excel 数据 ===", "HIGHLIGHT")
            df = pd.read_excel(self.excel_path, dtype=str).fillna('')
            df.columns = [str(c).strip() for c in df.columns]

            name_col = next((col for col in df.columns if '姓名' in col), None)
            xuhao_col = next((col for col in df.columns if '序号' in col), None)
            emp_id_col = next((col for col in df.columns if any(k in col for k in ['工号', '编号', '员工编码'])), None)

            if not name_col:
                self.log("❌ [错误] 未能在 Excel 中找到包含“姓名”字样的列！", "ERROR")
                return False

            self.log(f"  └─ 姓名列 -> [{name_col}] | 序号列 -> [{xuhao_col or '未找到'}] | 工号列 -> [{emp_id_col or '未找到'}]", "INFO")
            self.log(f"  └─ 图片类型共 {len(self.image_types)} 种: {', '.join(t['name'] for t in self.image_types)}", "INFO")

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
                emp_id = context.get(emp_id_col, '') if emp_id_col else ''
                if not name:
                    continue

                self.log(f"[{index + 1}/{total}] 正在生成: {name}", "INFO")
                tpl = DocxTemplate(self.template_path)

                # 智能匹配：同一目录下按关键词区分类型，其余兜底
                type_images = match_types_for_person(self.image_types, name, emp_id)

                for type_info in self.image_types:
                    type_name = type_info["name"]
                    imgs = type_images.get(type_name, [])
                    if imgs:
                        sub = tpl.new_subdoc()
                        inserted = 0
                        for img_path in imgs:
                            buf = prepare_image(img_path)
                            if buf:
                                sub.add_paragraph().add_run().add_picture(buf, width=Cm(15))
                                inserted += 1
                            else:
                                self.log(f"  └─ [{type_name}] 跳过无法识别的图片: {os.path.basename(img_path)}", "WARNING")
                        if inserted > 0:
                            context[type_name] = sub
                            self.log(f"  └─ [{type_name}] 匹配到 {inserted} 张", "SUCCESS")
                        else:
                            context[type_name] = f"【未找到{type_name}图片】"
                            self.log(f"  └─ [{type_name}] 全部图片无法识别", "WARNING")
                    else:
                        context[type_name] = f"【未找到{type_name}图片】"
                        self.log(f"  └─ [{type_name}] 未找到", "WARNING")

                tpl.render(context)
                xuhao = context.get(xuhao_col, '') if xuhao_col else str(index + 1)
                if not xuhao:
                    xuhao = str(index + 1)
                out_file = os.path.join(self.output_dir, f"{xuhao}_{name}_证明.docx")
                try:
                    tpl.save(out_file)
                    generated_files.append(out_file)
                    success_count += 1
                except PermissionError:
                    self.log(f"  └─ ⚠️ 文件被占用，已跳过: {name}", "WARNING")

            self.log("\n=== [阶段3] 汇总文档并回写 Excel 索引 ===", "HIGHLIGHT")
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            merged_out = os.path.join(self.output_dir, f"00_合并汇总_人员证明_{timestamp}.docx")
            if success_count > 0:
                self.log(f"  └─ 开始无缝合并 {success_count} 份独立文档...", "INFO")
                try:
                    master = Document(generated_files[0])
                    composer = Composer(master)
                    for file_path in generated_files[1:]:
                        composer.append(Document(file_path))
                    composer.save(merged_out)
                    self.log(f"  └─ 汇总文档生成成功！", "SUCCESS")
                except Exception as e:
                    self.log(f"  └─ [报错] 合并过程发生异常: {str(e)}", "ERROR")

            self._writeback_excel(name_col, xuhao_col)

            self.log(f"\n=== 🏁 任务结束 (共处理 {success_count} 人) ===", "HIGHLIGHT")
            return merged_out if success_count > 0 else None

        except Exception as e:
            err = str(e).strip() or f"{type(e).__name__}"
            self.log(f"❌ 核心引擎发生严重异常: {err}", "ERROR")
            import traceback
            traceback.print_exc()
            return None

    def _writeback_excel(self, name_col, xuhao_col):
        """回写 Excel：从 F 列起为每个图片类型生成一列目录索引，自动适配类型数量。"""
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
                return

            start_col = 6  # 从 F 列开始
            for type_idx, type_info in enumerate(self.image_types):
                col = start_col + type_idx
                type_name = type_info["name"]
                ws.cell(row=1, column=col).value = f"{type_name}目录"
                for r in range(2, ws.max_row + 1):
                    cell_name = ws.cell(row=r, column=name_col_idx).value
                    cell_xuhao = ws.cell(row=r, column=xuhao_col_idx).value if xuhao_col_idx else r - 1
                    if not cell_name or not str(cell_name).strip():
                        continue
                    ws.cell(row=r, column=col).value = f"{cell_xuhao}.{type_idx + 1} {cell_name}-{type_name}证明"

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            updated_excel_path = os.path.join(self.output_dir, f"00_人员清单_已自动更新目录_{timestamp}.xlsx")
            wb.save(updated_excel_path)
            self.log(f"  └─ 目录索引已填充 {len(self.image_types)} 个类型列 (从 F 列起)！", "SUCCESS")
        except Exception as e:
            self.log(f"  └─ [警告] 更新 Excel 发生异常 (文件可能被占用): {str(e)}", "WARNING")


# ==========================================
# 2. 界面交互类 (UI) - 动态图片类型列表 + 关键词识别
# ==========================================
class BatchImageInsertUI(BaseEngineUI):
    def __init__(self, parent_frame):
        super().__init__(parent_frame)
        self.latest_merged_path = ""
        self.type_rows = []
        self.setup_ui()
        self._load_configured_types()

    def setup_ui(self):
        self.parent.columnconfigure(1, weight=1)
        self.parent.rowconfigure(4, weight=1)

        ttk.Label(self.parent, text="人员清单(Excel):").grid(row=0, column=0, padx=10, pady=8, sticky="e")
        self.excel_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.excel_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_excel).grid(row=0, column=2, padx=10, pady=8)

        ttk.Label(self.parent, text="Word模板(.docx):").grid(row=1, column=0, padx=10, pady=8, sticky="e")
        self.tpl_entry = ttk.Entry(self.parent, font=("Microsoft YaHei UI", 10))
        self.tpl_entry.grid(row=1, column=1, padx=5, pady=8, sticky="ew")
        ttk.Button(self.parent, text="浏览...", command=self.select_tpl).grid(row=1, column=2, padx=10, pady=8)

        # --- 图片类型配置（动态增删 + 关键词识别） ---
        types_frame_outer = ttk.LabelFrame(self.parent, text=" 图片类型配置（名称 + 目录 + 识别关键词，模板用 {{类型名}} 占位） ")
        types_frame_outer.grid(row=2, column=0, columnspan=3, padx=10, pady=8, sticky="ew")
        types_frame_outer.columnconfigure(0, weight=1)

        btn_row = ttk.Frame(types_frame_outer)
        btn_row.grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        ttk.Button(btn_row, text="➕ 添加类型", command=self._on_add_type).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="📄 从模板识别类型", command=self._on_scan_template).pack(side=tk.LEFT)

        self.types_frame = ttk.Frame(types_frame_outer)
        self.types_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(2, 8))
        self.types_frame.columnconfigure(1, weight=1)

        ttk.Label(self.types_frame, text="类型名称", font=("Microsoft YaHei UI", 9, "bold")).grid(row=0, column=0, padx=(0, 5), sticky="w")
        ttk.Label(self.types_frame, text="图片目录", font=("Microsoft YaHei UI", 9, "bold")).grid(row=0, column=1, padx=(0, 5), sticky="w")
        ttk.Label(self.types_frame, text="识别关键词(逗号分隔,留空=其余图片)", font=("Microsoft YaHei UI", 9, "bold")).grid(row=0, column=2, padx=(0, 5), sticky="w")

        # --- 执行按钮 ---
        self.execute_btn = tk.Button(self.parent, text="▶ 全自动合图并填表", bg="#4CAF50", fg="white",
                                     font=("Microsoft YaHei UI", 11, "bold"), command=self.on_execute)
        self.execute_btn.grid(row=3, column=0, columnspan=3, pady=12, sticky=tk.EW, padx=20, ipady=5)

        # --- 日志区 ---
        tk.Label(self.parent, text="运行追踪与日志:", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333").grid(row=4, column=0, sticky=tk.W, padx=15, pady=(8, 5))
        self.log_text = scrolledtext.ScrolledText(
            self.parent, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
            padx=12, pady=12, relief=tk.FLAT, state='disabled'
        )
        self.log_text.grid(row=5, column=0, columnspan=3, padx=15, pady=5, sticky=tk.NSEW)

        self.log_text.tag_config("INFO", foreground="#2C3E50")
        self.log_text.tag_config("SUCCESS", foreground="#198754", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("WARNING", foreground="#D35400", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("ERROR", foreground="#DC3545", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_text.tag_config("HIGHLIGHT", foreground="#8E44AD", font=("Microsoft YaHei UI", 11, "bold"))
        self.log_text.tag_config("LINK", foreground="#0078D7", underline=True, font=("Microsoft YaHei UI", 11, "bold"))

        self.log_text.tag_bind("LINK", "<Button-1>", self.open_generated_file)
        self.log_text.tag_bind("LINK", "<Enter>", lambda e: self.log_text.config(cursor="hand2"))
        self.log_text.tag_bind("LINK", "<Leave>", lambda e: self.log_text.config(cursor=""))

        self.log_to_ui("✨ 智能合图填表引擎已就绪（支持关键词识别图片类型）！", "HIGHLIGHT")

    # ---------- 类型行增删 ----------
    def _add_type_row(self, name="", dir="", keywords=""):
        row_idx = len(self.type_rows) + 1
        name_var = tk.StringVar(value=name)
        dir_var = tk.StringVar(value=dir)
        kw_var = tk.StringVar(value=keywords)

        name_entry = ttk.Entry(self.types_frame, textvariable=name_var, width=12, font=("Microsoft YaHei UI", 10))
        name_entry.grid(row=row_idx, column=0, padx=(0, 5), pady=2, sticky="w")

        dir_entry = ttk.Entry(self.types_frame, textvariable=dir_var, font=("Microsoft YaHei UI", 10))
        dir_entry.grid(row=row_idx, column=1, padx=(0, 5), pady=2, sticky="ew")

        kw_entry = ttk.Entry(self.types_frame, textvariable=kw_var, width=22, font=("Microsoft YaHei UI", 10))
        kw_entry.grid(row=row_idx, column=2, padx=(0, 5), pady=2, sticky="ew")

        browse_btn = ttk.Button(self.types_frame, text="浏览", width=6, command=lambda v=dir_var: self._browse_dir(v))
        browse_btn.grid(row=row_idx, column=3, padx=(0, 4), pady=2)

        delete_btn = ttk.Button(self.types_frame, text="✖", width=3)
        delete_btn.grid(row=row_idx, column=4, pady=2)

        row_data = {"name_var": name_var, "dir_var": dir_var, "kw_var": kw_var}
        delete_btn.config(command=lambda r=row_data: self._on_remove_type(r))
        row_data.update({
            "name_entry": name_entry, "dir_entry": dir_entry, "kw_entry": kw_entry,
            "browse_btn": browse_btn, "delete_btn": delete_btn, "row": row_idx,
        })
        self.type_rows.append(row_data)
        return row_data

    def _remove_type_row(self, row_data):
        if row_data in self.type_rows:
            self.type_rows.remove(row_data)
        for w in ("name_entry", "dir_entry", "kw_entry", "browse_btn", "delete_btn"):
            widget = row_data.get(w)
            if widget is not None:
                widget.destroy()
        self._relayout_types()

    def _relayout_types(self):
        """重新排列所有类型行的 grid 行号（从 1 起连续编号），避免删除后行号冲突。"""
        for idx, r in enumerate(self.type_rows, start=1):
            r["name_entry"].grid(row=idx, column=0)
            r["dir_entry"].grid(row=idx, column=1)
            r["kw_entry"].grid(row=idx, column=2)
            r["browse_btn"].grid(row=idx, column=3)
            r["delete_btn"].grid(row=idx, column=4)
            r["row"] = idx

    def _collect_types(self):
        types = []
        for r in self.type_rows:
            name = r["name_var"].get().strip()
            dir_ = r["dir_var"].get().strip()
            keywords = r["kw_var"].get().strip()
            if name:
                types.append({"name": name, "dir": dir_, "keywords": keywords})
        return types

    def _persist_types(self):
        save_image_types(self._collect_types())

    def _load_configured_types(self):
        for t in load_image_types():
            self._add_type_row(t["name"], t["dir"], t.get("keywords", ""))

    def _browse_dir(self, dir_var):
        path = filedialog.askdirectory(title="选择图片目录")
        if path:
            dir_var.set(path)

    def _on_add_type(self):
        self._add_type_row()
        self._persist_types()

    def _on_remove_type(self, row_data):
        self._remove_type_row(row_data)
        self._persist_types()

    def _on_scan_template(self):
        """扫描模板占位符，智能识别图片类型（排除 Excel 文本字段 + 关键词过滤）。"""
        tpl_path = self.tpl_entry.get().strip()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showwarning("提示", "请先选择有效的 Word 模板！")
            return
        variables = extract_template_variables(tpl_path)
        excel_path = self.excel_entry.get().strip()
        excel_columns = get_excel_columns(excel_path) if excel_path else []
        image_vars = identify_image_type_variables(variables, excel_columns)
        existing = {t["name"] for t in self._collect_types()}
        added = []
        for v in sorted(image_vars):
            if v and v not in existing:
                self._add_type_row(name=v)
                existing.add(v)
                added.append(v)
        if added:
            self._persist_types()
            hint = "（请补充目录与关键词）"
            if not excel_path:
                hint += "；建议先选择 Excel 清单以更精确排除文本字段"
            self.log_to_ui(f"  └─ 智能识别并新增 {len(added)} 个图片类型: {', '.join(added)}{hint}", "SUCCESS")
        else:
            self.log_to_ui("  └─ 未识别到新的图片类型占位符（文本字段已自动排除）", "INFO")

    # ---------- 文件选择与调度 ----------
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

    def on_execute(self):
        excel_path = self.excel_entry.get().strip()
        tpl_path = self.tpl_entry.get().strip()
        image_types = self._collect_types()

        if not excel_path or not tpl_path:
            messagebox.showwarning("参数缺失", "请先填写 Excel 清单与 Word 模板路径！")
            return
        if not image_types:
            messagebox.showwarning("参数缺失", "请至少配置一个图片类型（名称 + 图片目录）！")
            return

        self._persist_types()

        output_dir = os.path.join(os.path.dirname(tpl_path), "已生成_人员证明")

        self.execute_btn.config(state="disabled", text="引擎高速运转中...", bg="#E9ECEF", fg="#ADB5BD")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        self.log_to_ui("🚀 启动智能合图填表流水线...", "HIGHLIGHT")

        engine = BatchImageInsertEngine(
            excel_path=excel_path,
            template_path=tpl_path,
            image_types=image_types,
            output_dir=output_dir,
            log_callback=self.log_to_ui
        )

        def run_thread():
            merged_file_path = engine.execute()

            def restore_btn():
                self.execute_btn.config(state="normal", text="▶ 全自动合图并填表", bg="#4CAF50", fg="white")

            self.parent.after(0, restore_btn)

            if merged_file_path:
                file_name = os.path.basename(merged_file_path)
                self.log_success_link("★ 任务全线竣工！点击立即打开验收 ➔ ", file_name, merged_file_path)

        threading.Thread(target=run_thread, daemon=True).start()
