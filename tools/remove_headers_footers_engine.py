import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

from tkinterdnd2 import DND_FILES, TkinterDnD

from utils.tree_ui_base import TreeEngineUI


# ==========================================
# 核心：清除 docx 文档的页眉页脚
# ==========================================
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def remove_headers_footers(src, dst):
    """清除 docx 文档中所有节的页眉、页脚，写入 dst（不覆盖原件）。

    实现原理：docx 本质是 zip 包，直接操作其 XML——
    1. 移除 document.xml 中所有节的 headerReference / footerReference / titlePg 标记；
    2. 移除 document.xml.rels 中指向页眉/页脚部件的关系；
    3. 移除包内残留的 header*.xml / footer*.xml 部件及其关系文件；
    4. 清理 [Content_Types].xml 中对应声明；
    5. 清除 settings.xml 中的奇偶页页眉页脚设置。
    """
    if not os.path.isfile(src):
        raise FileNotFoundError(f"找不到原文件：{src}")
    if os.path.abspath(src) == os.path.abspath(dst):
        raise ValueError("输出路径不能与原文件相同。")

    with ZipFile(src, "r") as zin:
        files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    # 1. 移除节中的页眉、页脚引用与「首页不同」标记
    doc_path = "word/document.xml"
    doc_root = ET.fromstring(files[doc_path])
    for node in doc_root.iter():
        if node.tag == f"{{{W}}}sectPr":
            for child in list(node):
                if child.tag in (
                    f"{{{W}}}headerReference",
                    f"{{{W}}}footerReference",
                    f"{{{W}}}titlePg",
                ):
                    node.remove(child)
    files[doc_path] = ET.tostring(doc_root, encoding="utf-8", xml_declaration=True)

    # 2. 移除指向页眉/页脚部件的关系
    rels_path = "word/_rels/document.xml.rels"
    rels_root = ET.fromstring(files[rels_path])
    removed_parts = set()
    for rel in list(rels_root):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/header") or rel_type.endswith("/footer"):
            target = rel.attrib.get("Target", "")
            if target.startswith("/"):
                part = target.lstrip("/")
            else:
                part = str(PurePosixPath("word") / PurePosixPath(target))
            part = str(PurePosixPath(part))
            removed_parts.add(part)
            rels_root.remove(rel)
    files[rels_path] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    # 3. 移除包内残留的页眉/页脚 XML 部件及它们的关系文件
    for name in list(files):
        path = PurePosixPath(name)
        if (
            path.parent == PurePosixPath("word")
            and path.name.startswith(("header", "footer"))
            and path.suffix == ".xml"
        ):
            removed_parts.add(name)

    for part in removed_parts:
        files.pop(part, None)
        part_path = PurePosixPath(part)
        part_rels = str(part_path.parent / "_rels" / f"{part_path.name}.rels")
        files.pop(part_rels, None)

    # 4. 清理内容类型清单中对应部件的声明
    ct_path = "[Content_Types].xml"
    ct_root = ET.fromstring(files[ct_path])
    removed_names = {"/" + part for part in removed_parts}
    for item in list(ct_root):
        if (
            item.tag == f"{{{CT}}}Override"
            and item.attrib.get("PartName") in removed_names
        ):
            ct_root.remove(item)
    files[ct_path] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

    # 5. 清除奇偶页页眉页脚设置
    settings_path = "word/settings.xml"
    if settings_path in files:
        settings_root = ET.fromstring(files[settings_path])
        for parent in settings_root.iter():
            for child in list(parent):
                if child.tag == f"{{{W}}}evenAndOddHeaders":
                    parent.remove(child)
        files[settings_path] = ET.tostring(settings_root, encoding="utf-8", xml_declaration=True)

    # 写入新文件
    with ZipFile(dst, "w", ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


# ==========================================
# 界面交互类 (UI)
# ==========================================
class RemoveHeadersFootersUI(TreeEngineUI):
    """清除 Word 文档页眉页脚引擎（支持拖拽、文件/文件夹批量处理）。"""

    COLOR_START = "#8E44AD"
    START_TEXT = "▶ 开始清除页眉页脚"
    RESUME_TEXT = "▶ 继续清除"
    ANIM_FRAMES = ("⚙ 清除中.  ", "⚙ 清除中.. ", "⚙ 清除中...")
    CANCEL_CONFIRM = "确定要中止当前的清除任务吗？\n已处理完成的文档将被保留。"
    STOPPING_STATUS_TEXT = "正在安全停止任务..."
    PAUSE_STATUS_TEXT = "任务已挂起，您可以随时删除待办列表中的文档..."
    RESUME_STATUS_TEXT = "任务已恢复，正在继续清除..."
    DRAG_RUNNING_WARN = "任务正在运行，请先暂停后再拖入新文档！"
    DROP_EMPTY_MSG = "拖入的项目中没有找到有效的 Word 文档！"
    ADD_FILES_TITLE = "选择 Word 文档"
    ADD_FILES_FILETYPES = [("Word 文档", "*.docx")]
    ADD_FOLDER_TITLE = "选择包含 Word 的文件夹"
    ADD_FOLDER_EMPTY_MSG = "该目录下未找到有效的 Word 文件！"
    ITEM_UNIT = "份"
    ITEM_LABEL = "文档"

    def __init__(self, parent_frame):
        super().__init__(parent_frame)
        self.target_var = tk.StringVar()
        self.setup_ui()

        # 注册拖拽事件支持
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.on_drag_drop)

    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

        # 1. 源文档添加区
        tk.Label(frame, text="1. 添加源文档:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)

        btn_source_frame = tk.Frame(frame)
        btn_source_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)

        tk.Button(btn_source_frame, text="📄 添加文件", font=("Microsoft YaHei UI", 9),
                  command=self.add_files).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="📁 添加文件夹 (扫描子目录)", font=("Microsoft YaHei UI", 9),
                  command=self.add_folder).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="▤ 清空列表", font=("Microsoft YaHei UI", 9),
                  command=self.clear_list).pack(side=tk.LEFT)

        tk.Label(btn_source_frame, text="💡", font=("Microsoft YaHei UI", 12), fg="#F39C12").pack(side=tk.LEFT, padx=(15, 0))
        tk.Label(btn_source_frame, text="提示：支持将 Word 文档和文件夹直接拖拽到下方列表。",
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        # 2. 输出目录选择
        tk.Label(frame, text="2. 保存至:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10)).grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.select_target, width=12).grid(row=1, column=2, pady=5)

        # 3. 控制按钮区
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)

        self.btn_pause = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 4. 可视化数据表格区
        list_frame = tk.Frame(frame)
        list_frame.grid(row=3, column=0, columnspan=3, sticky=tk.NSEW, pady=5)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        columns = ("id", "filename", "target_type", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)

        self.tree.heading("id", text="序号")
        self.tree.column("id", width=60, anchor="center")
        self.tree.heading("filename", text="文档名称")
        self.tree.column("filename", width=400, anchor="w")
        self.tree.heading("target_type", text="目标格式")
        self.tree.column("target_type", width=100, anchor="center")
        self.tree.heading("status", text="处理状态")
        self.tree.column("status", width=250, anchor="center")

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)

        self.tree.tag_configure("success", foreground="#198754")
        self.tree.tag_configure("error", foreground="#DC3545")
        self.tree.tag_configure("processing", foreground="#0078D7")
        self.tree.tag_configure("pending", foreground="#6C757D")
        self.tree.tag_configure("warning", foreground="#D35400")

        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Delete>", self.delete_selected_items)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.context_menu = tk.Menu(self.tree, tearoff=0, font=("Microsoft YaHei UI", 10))
        self.context_menu.add_command(label="✖ 移除选中文档 (Del)", command=self.delete_selected_items)

        # 5. 底部汇总栏
        bottom_frame = tk.Frame(frame)
        bottom_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(10, 0))

        self.lbl_icon_var = tk.StringVar(value="💡")
        self.lbl_icon = tk.Label(bottom_frame, textvariable=self.lbl_icon_var, font=("Microsoft YaHei UI", 12))
        self.lbl_icon.pack(side=tk.LEFT)

        self.lbl_summary_var = tk.StringVar(value="提示：双击表格记录可打开文档；按 Delete 键可踢除文档。")
        self.lbl_summary = tk.Label(bottom_frame, textvariable=self.lbl_summary_var, font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_summary.pack(side=tk.LEFT)

        self.btn_open_dir = tk.Button(bottom_frame, text="📂 打开输出目录", font=("Microsoft YaHei UI", 9),
                                      command=self.open_output_dir, state=tk.DISABLED)
        self.btn_open_dir.pack(side=tk.RIGHT)

        self.update_button_ui("idle")
        self.set_status("💡", "提示：双击表格记录可打开文档；按 Delete 键可踢除文档。", "#F39C12", "#2C3E50")

    def is_valid_file(self, path):
        return path.lower().endswith('.docx')

    def build_row_values(self, idx, path):
        return (idx, os.path.basename(path), ".docx", "⌛ 等待中")

    def _get_safe_path(self, target_dir, base_name):
        """生成不覆盖已有文件的输出路径（同名时自动加序号，如 xxx(1).docx）。"""
        stem = f"{base_name}_已清除页眉页脚"
        out_path = os.path.join(target_dir, f"{stem}.docx")
        counter = 1
        while os.path.exists(out_path):
            out_path = os.path.join(target_dir, f"{stem}({counter}).docx")
            counter += 1
        return out_path

    def start_conversion_thread(self):
        tgt = self.target_var.get().strip().strip('"').strip("'")

        if not tgt:
            messagebox.showwarning("提示", "请先选择输出保存位置！")
            return

        pending_items = list(self.tree.get_children())
        if not pending_items:
            messagebox.showinfo("提示", "列表中没有待处理的文档！可以点击上方添加或直接拖入文件。")
            return

        # 输出文件自动重命名，无需「目录非空」防覆盖确认
        self.item_outpath_map.clear()
        for item_id in pending_items:
            vals = list(self.tree.item(item_id, "values"))
            vals[3] = "⌛ 等待中"
            self.tree.item(item_id, values=vals, tags=("pending",))

        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        self.update_button_ui("running")
        self.btn_open_dir.config(state=tk.DISABLED)

        self.set_status("⚙", f"正在清除页眉页脚，队列中共 {len(pending_items)} 份文档...", "#0078D7")

        threading.Thread(target=self.process_conversion, args=(pending_items, tgt), daemon=True).start()

    # ================= 核心处理逻辑 =================
    def process_conversion(self, pending_items, tgt):
        success_count = 0
        error_count = 0

        try:
            for item_id in pending_items:
                if self._stop_event.is_set():
                    break
                self._pause_event.wait()

                if not self.tree.exists(item_id) or item_id not in self.item_filepath_map:
                    continue

                file_path = self.item_filepath_map[item_id]
                file_name = os.path.basename(file_path)
                idx = self.tree.item(item_id, "values")[0]

                self.update_tree_row(item_id, (idx, file_name, ".docx", "🧹 正在清除..."), "processing")

                try:
                    if not os.path.exists(tgt):
                        os.makedirs(tgt)

                    base_name = os.path.splitext(file_name)[0]
                    out_path = self._get_safe_path(tgt, base_name)

                    remove_headers_footers(file_path, out_path)

                    self.item_outpath_map[item_id] = out_path
                    self.update_tree_row(item_id, (idx, file_name, ".docx", "✔ 清除成功"), "success")
                    success_count += 1

                except PermissionError:
                    self.update_tree_row(item_id, (idx, file_name, ".docx", "✖ 文件被占用，请关闭后重试"), "error")
                    error_count += 1
                except Exception as e:
                    self.update_tree_row(item_id, (idx, file_name, ".docx", f"✖ {str(e)[:20]}"), "error")
                    error_count += 1

            if self._stop_event.is_set():
                for item_id in pending_items:
                    if self.tree.exists(item_id):
                        tags = self.tree.item(item_id, "tags")
                        if "pending" in tags:
                            vals = list(self.tree.item(item_id, "values"))
                            vals[3] = "✖ 已取消"
                            self.update_tree_row(item_id, vals, "error", auto_scroll=False)

            def finish_ui():
                self.is_running = False
                self.update_button_ui("idle")
                self.btn_open_dir.config(state=tk.NORMAL)
                if self._stop_event.is_set():
                    self.set_status("■", "任务已中断。", "#DC3545")
                else:
                    self.set_status("✔", f"任务完成！成功 {success_count} 份，失败 {error_count} 份。", "#198754")

            self.parent.after(0, finish_ui)

        except Exception as global_e:
            error_msg = str(global_e)
            self.parent.after(0, lambda e=error_msg: self.set_status("✖", f"发生全局错误: {e}", "#DC3545"))
            self.parent.after(0, lambda: self.update_button_ui("idle"))
