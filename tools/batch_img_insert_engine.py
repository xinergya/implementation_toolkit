import os
import re
import io
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from collections import OrderedDict
import pandas as pd
import openpyxl
from datetime import datetime
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
from docx.shared import Cm

from utils.ui_base import BaseEngineUI


# ==========================================
# 图片类型配置（动态可扩展）
# ==========================================
def _get_config_path():
    """图片类型配置文件路径（优先 %APPDATA%，回退脚本目录）。"""
    appdata = os.environ.get("APPDATA")
    if appdata:
        return os.path.join(appdata, "EfficiencyToolbox", "img_type_config.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "img_type_config.json")


# 默认类型示例：keywords 为空表示「兜底类型」，接受该目录下未被其他类型匹配的图片
DEFAULT_IMAGE_TYPES = [
    {"name": "身份证", "dir": "", "keywords": "身份证"},
    {"name": "学历证书", "dir": "", "keywords": "毕业,学位,学士"},
    {"name": "资质证书", "dir": "", "keywords": ""},
    {"name": "社保证明", "dir": "", "keywords": ""},
]


def load_image_types():
    """加载图片类型配置；文件缺失或解析失败时回退默认类型。兼容旧配置（无 keywords 字段）。"""
    try:
        with open(_get_config_path(), 'r', encoding='utf-8') as f:
            data = json.load(f)
        types = data.get("image_types", [])
        if types:
            result = []
            for t in types:
                result.append({
                    "name": str(t.get("name", "")),
                    "dir": str(t.get("dir", "")),
                    "keywords": str(t.get("keywords", "")),
                })
            return result
    except Exception:
        pass
    return [dict(t) for t in DEFAULT_IMAGE_TYPES]


def save_image_types(types):
    """保存图片类型配置（名称 + 目录 + 识别关键词）。"""
    try:
        path = _get_config_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({"image_types": types}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存图片类型配置失败：{e}")


def extract_template_variables(template_path):
    """扫描 Word 模板，提取所有 {{变量}} 占位符变量名。"""
    variables = set()
    if not template_path or not os.path.exists(template_path):
        return variables
    try:
        doc = Document(template_path)
        paras = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    paras.extend(cell.paragraphs)
        pattern = re.compile(r'{{\s*([A-Za-z0-9_\u4e00-\u9fa5]+)\s*}}')
        for para in paras:
            for m in pattern.finditer(para.text):
                variables.add(m.group(1))
    except Exception:
        pass
    return variables


def parse_keywords(keywords_str):
    """把逗号分隔的关键词字符串解析为列表（去空、去重）。"""
    seen = []
    for k in (keywords_str or "").replace('，', ',').split(','):
        k = k.strip()
        if k and k not in seen:
            seen.append(k)
    return seen


def match_keywords(img_path, keywords):
    """判断图片文件名是否含任一关键词（不区分大小写）。"""
    fname = os.path.basename(img_path).lower()
    return any(kw.lower() in fname for kw in keywords)


# ==========================================
# 1. 核心处理引擎类 (Engine)
# ==========================================
class BatchImageInsertEngine:
    VALID_EXTS = ('.jpg', '.jpeg', '.png', '.bmp', '.jfif', '.gif')

    def __init__(self, excel_path, template_path, image_types, output_dir, log_callback=None):
        self.excel_path = excel_path
        self.template_path = template_path
        self.image_types = [t for t in image_types if t.get("name")]
        self.output_dir = output_dir
        self.log = log_callback if log_callback else (lambda msg, level="INFO", **kwargs: print(f"[{level}] {msg}"))

    def find_person_images(self, name, emp_id, search_dir):
        """查找某人员的所有候选图片，兼容「分人子目录」与「平铺」两种结构。

        1. 若 search_dir 下存在以姓名/工号命名的子文件夹，则收集这些子文件夹内的全部图片；
        2. 否则视为平铺目录，收集文件名含姓名/工号的图片。
        """
        images = []
        if not search_dir or not os.path.exists(search_dir):
            return images
        try:
            entries = os.listdir(search_dir)
        except Exception:
            return images

        # 1. 分人子目录模式
        matched_subdirs = []
        for entry in entries:
            full = os.path.join(search_dir, entry)
            if os.path.isdir(full) and (name in entry or (emp_id and emp_id in entry)):
                matched_subdirs.append(full)
        if matched_subdirs:
            for sd in matched_subdirs:
                for root, dirs, files in os.walk(sd):
                    for f in files:
                        if f.lower().endswith(self.VALID_EXTS):
                            images.append(os.path.join(root, f))
            return images

        # 2. 平铺模式
        for f in entries:
            if f.lower().endswith(self.VALID_EXTS) and (name in f or (emp_id and emp_id in f)):
                images.append(os.path.join(search_dir, f))
        return images

    @staticmethod
    def _prepare_image(img_path):
        """用 Pillow 规范化图片为 python-docx 可识别的标准格式，返回 BytesIO。

        解决三类问题：① 某些 JPEG 缺少 JFIF/Exif 头（文件头为 ffd8ffdb），python-docx
        无法识别；② 源图为 WebP/HEIC 等 python-docx 不支持的格式；③ 按 EXIF 方向信息
        转正图片，避免手机拍摄的照片在文档中横向/倒置摆放。转换失败返回 None。
        """
        try:
            from PIL import Image, ImageOps
            with Image.open(img_path) as im:
                fmt = im.format
                # 应用 EXIF 方向信息，把图片转正（无 EXIF 或已正向则原样返回）
                im = ImageOps.exif_transpose(im)
                buf = io.BytesIO()
                has_alpha = im.mode in ('RGBA', 'LA', 'P')
                if fmt == 'JPEG':
                    im.convert('RGB').save(buf, format='JPEG', quality=92)
                elif has_alpha:
                    im.convert('RGBA').save(buf, format='PNG')
                else:
                    im.convert('RGB').save(buf, format='PNG')
                buf.seek(0)
                return buf
        except Exception:
            return None

    def _match_types_for_person(self, name, emp_id):
        """对某人员，按目录分组执行「关键词优先 + 兜底」匹配，返回 {类型名: [图片路径]}。"""
        dir_groups = OrderedDict()
        for ti in self.image_types:
            dir_groups.setdefault(ti["dir"], []).append(ti)

        type_images = {}
        for type_dir, types in dir_groups.items():
            person_imgs = self.find_person_images(name, emp_id, type_dir)
            used = set()
            # 第一遍：有关键词的类型按文件名匹配
            for ti in types:
                kws = parse_keywords(ti.get("keywords", ""))
                if kws:
                    matched = [p for p in person_imgs if p not in used and match_keywords(p, kws)]
                    type_images[ti["name"]] = matched
                    used.update(matched)
            # 第二遍：关键词为空的类型接受剩余图片（兜底）
            remaining = [p for p in person_imgs if p not in used]
            for ti in types:
                if not parse_keywords(ti.get("keywords", "")):
                    type_images[ti["name"]] = remaining
        return type_images

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
                type_images = self._match_types_for_person(name, emp_id)

                for type_info in self.image_types:
                    type_name = type_info["name"]
                    imgs = type_images.get(type_name, [])
                    if imgs:
                        sub = tpl.new_subdoc()
                        inserted = 0
                        for img_path in imgs:
                            buf = self._prepare_image(img_path)
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
        """重新排列所有类型行的 grid 行号（从 1 起连续编号）。

        删除中间行后若不重排，`len(type_rows)+1` 计算的新行号会与已占用的
        空行冲突，导致新行覆盖最后一行。此方法保证行号始终连续。
        """
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
        """扫描模板占位符，自动识别并新增图片类型（目录/关键词留空待填）。"""
        tpl_path = self.tpl_entry.get().strip()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showwarning("提示", "请先选择有效的 Word 模板！")
            return
        variables = extract_template_variables(tpl_path)
        existing = {t["name"] for t in self._collect_types()}
        added = []
        for v in sorted(variables):
            if v and v not in existing:
                self._add_type_row(name=v)
                existing.add(v)
                added.append(v)
        if added:
            self._persist_types()
            self.log_to_ui(f"  └─ 从模板识别并新增 {len(added)} 个图片类型: {', '.join(added)}（请补充目录与关键词）", "SUCCESS")
        else:
            self.log_to_ui("  └─ 模板中未发现新的图片类型占位符", "INFO")

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
