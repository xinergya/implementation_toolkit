import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 核心库：Word 控制与 PDF 高清渲染
import win32com.client
import pythoncom
import fitz  # PyMuPDF
# 核心拖拽支持库
from tkinterdnd2 import DND_FILES, TkinterDnD

from utils.tree_ui_base import TreeEngineUI
from utils.constants import com_error_kind
from utils.pathutil import is_valid_word_file


class WordToImageUI(TreeEngineUI):
    """Word 转 高清图片 引擎 (支持混合拖拽 + 动态反馈 + 极简输出)"""

    COLOR_START = "#2980B9"
    START_TEXT = "▶ 开始排版渲染"
    RESUME_TEXT = "▶ 继续渲染"
    ANIM_FRAMES = ("⚙ 解析排版中.  ", "⚙ 解析排版中.. ", "⚙ 解析排版中...")
    CANCEL_CONFIRM = "确定要中止当前的渲染任务吗？\n已输出完成的图片将被保留。"
    STOPPING_STATUS_TEXT = "正在安全停止任务，清理底层临时缓存..."
    PAUSE_STATUS_TEXT = "任务已挂起，您可以随时删除待办列表中的文档..."
    RESUME_STATUS_TEXT = "任务已恢复，正在全力排版渲染..."
    DRAG_RUNNING_WARN = "任务正在运行，请先暂停后再拖入新文档！"
    DROP_EMPTY_MSG = "拖入的项目中没有找到有效的 Word 文档！"
    ADD_FILES_TITLE = "选择 Word 文档"
    ADD_FILES_FILETYPES = [("Word 文档", "*.doc;*.docx")]
    ADD_FOLDER_TITLE = "选择包含 Word 的文件夹"
    ADD_FOLDER_EMPTY_MSG = "该目录下未找到有效的 Word 文件！"

    def __init__(self, parent_frame):
        super().__init__(parent_frame)

        self.target_var = tk.StringVar()
        self.format_var = tk.StringVar(value="jpg")

        self.setup_ui()

        # 注册拖拽事件支持
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.on_drag_drop)

    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)  # 调整自适应行

        # 1. 源文档添加区 (合并文件与文件夹功能)
        tk.Label(frame, text="1. 添加源文档:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0,
                                                                                             sticky=tk.W, pady=5)

        btn_source_frame = tk.Frame(frame)
        btn_source_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)

        tk.Button(btn_source_frame, text="📄 添加文件", font=("Microsoft YaHei UI", 9),
                  command=self.add_files).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="📁 添加文件夹 (扫描子目录)", font=("Microsoft YaHei UI", 9),
                  command=self.add_folder).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="▤ 清空列表", font=("Microsoft YaHei UI", 9),
                  command=self.clear_list).pack(side=tk.LEFT)

        tk.Label(btn_source_frame, text="💡", font=("Microsoft YaHei UI", 12), fg="#F39C12").pack(side=tk.LEFT,
                                                                                                 padx=(15, 0))
        tk.Label(btn_source_frame, text="提示：支持将 Word 文件和文件夹直接混搭拖拽到下方列表中。",
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        # 2. 输出目录选择 (所见即所得，不追加相对路径)
        tk.Label(frame, text="2. 图片保存至:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0,
                                                                                             sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10)).grid(row=1, column=1,
                                                                                            sticky=tk.EW, padx=10,
                                                                                            pady=5, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.select_target, width=12).grid(row=1, column=2, pady=5)

        # 3. 图片格式与清晰度选择区
        tk.Label(frame, text="3. 渲染配置:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0, sticky=tk.W,
                                                                                           pady=5)
        config_frame = tk.Frame(frame)
        config_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=10, pady=5)

        tk.Radiobutton(config_frame, text="标准 JPG", variable=self.format_var, value="jpg",
                       font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=(0, 10))
        tk.Radiobutton(config_frame, text="高清 PNG", variable=self.format_var, value="png",
                       font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(config_frame, text="DPI清晰度:", font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=(10, 5))
        self.dpi_var = tk.StringVar(value="200 (超清)")
        self.dpi_combo = ttk.Combobox(config_frame, textvariable=self.dpi_var, font=("Microsoft YaHei UI", 10),
                                      width=20)
        self.dpi_combo['values'] = ("96 (标准)", "144 (高清)", "200 (超清)", "300 (印刷级)", "600 (典藏级)")
        self.dpi_combo.pack(side=tk.LEFT, ipady=3)
        tk.Label(config_frame, text="* 可下拉选择，或直接输入纯数字 (如: 120)", font=("Microsoft YaHei UI", 9),
                 fg="#6C757D").pack(side=tk.LEFT, padx=(10, 0))

        # 4. 控制按钮区 (状态机管理)
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)

        self.btn_pause = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 5. 可视化数据表格区
        list_frame = tk.Frame(frame)
        list_frame.grid(row=4, column=0, columnspan=3, sticky=tk.NSEW, pady=5)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        columns = ("id", "filename", "format", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)

        self.tree.heading("id", text="序号")
        self.tree.column("id", width=50, anchor="center")
        self.tree.heading("filename", text="文档名称")
        self.tree.column("filename", width=380, anchor="w")
        self.tree.heading("format", text="目标格式")
        self.tree.column("format", width=90, anchor="center")
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

        # 6. 底部汇总栏
        bottom_frame = tk.Frame(frame)
        bottom_frame.grid(row=5, column=0, columnspan=3, sticky=tk.EW, pady=(10, 0))

        self.lbl_icon_var = tk.StringVar(value="💡")
        self.lbl_icon = tk.Label(bottom_frame, textvariable=self.lbl_icon_var, font=("Microsoft YaHei UI", 12))
        self.lbl_icon.pack(side=tk.LEFT)

        self.lbl_summary_var = tk.StringVar(value="提示：双击记录可打开对应的图片文件夹；按 Delete 键可剔除文档。")
        self.lbl_summary = tk.Label(bottom_frame, textvariable=self.lbl_summary_var,
                                    font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_summary.pack(side=tk.LEFT)

        self.btn_open_dir = tk.Button(bottom_frame, text="📂 打开输出目录", font=("Microsoft YaHei UI", 9),
                                      command=self.open_output_dir, state=tk.DISABLED)
        self.btn_open_dir.pack(side=tk.RIGHT)

        self.update_button_ui("idle")
        self.set_status("💡", "提示：双击记录可打开对应的图片文件夹；按 Delete 键可剔除文档。", "#F39C12", "#2C3E50")

    def is_valid_file(self, path):
        return is_valid_word_file(path)

    def build_row_values(self, idx, path):
        return (idx, os.path.basename(path), f".{self.format_var.get().upper()}", "⌛ 等待中")

    def start_conversion_thread(self):
        tgt = self.target_var.get().strip().strip('"').strip("'")
        img_fmt = self.format_var.get()
        raw_dpi = self.dpi_var.get().strip()
        dpi_str = raw_dpi.split(' ')[0]

        try:
            dpi_value = int(dpi_str)
            if not (50 <= dpi_value <= 1200):
                raise ValueError
        except ValueError:
            messagebox.showwarning("参数错误", "请输入有效的 DPI 数值（建议 50 - 600 之间的纯数字）！")
            return

        if not tgt:
            messagebox.showwarning("提示", "请选择输出的保存位置！")
            return

        pending_items = list(self.tree.get_children())
        if not pending_items:
            messagebox.showinfo("提示", "列表中没有待处理的文档！可以点击上方添加或直接拖入文件。")
            return

        self.item_outpath_map.clear()
        for item_id in pending_items:
            vals = list(self.tree.item(item_id, "values"))
            vals[2] = f".{img_fmt.upper()}"
            vals[3] = "⌛ 等待中"
            self.tree.item(item_id, values=vals, tags=("pending",))

        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        self.update_button_ui("running")
        self.btn_open_dir.config(state=tk.DISABLED)

        threading.Thread(target=self.process_conversion, args=(pending_items, tgt, img_fmt, dpi_value),
                         daemon=True).start()

    # ================= 核心渲染与分发逻辑 =================
    def process_conversion(self, pending_items, tgt, img_fmt, dpi_value):
        pythoncom.CoInitialize()
        word_app = None
        success_count = 0
        error_count = 0
        seen_base_names = {}  # 用于处理同名文件输出目录冲突

        try:
            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False

            for item_id in pending_items:
                if self._stop_event.is_set():
                    break
                self._pause_event.wait()

                if not self.tree.exists(item_id) or item_id not in self.item_filepath_map:
                    continue

                file_path = self.item_filepath_map[item_id]
                file_name = os.path.basename(file_path)
                idx = self.tree.item(item_id, "values")[0]

                self.update_tree_row(item_id, (idx, file_name, f".{img_fmt.upper()}", "⚙ 启动底座排版中..."),
                                     "processing")

                doc = None
                temp_pdf_path = ""
                try:
                    base_name = os.path.splitext(file_name)[0]
                    # 同名文件冲突处理：输出目录名加序号，避免相互覆盖
                    if base_name in seen_base_names:
                        seen_base_names[base_name] += 1
                        doc_img_dir = os.path.join(tgt, f"{base_name}_{seen_base_names[base_name]}")
                    else:
                        seen_base_names[base_name] = 0
                        doc_img_dir = os.path.join(tgt, base_name)

                    if not os.path.exists(doc_img_dir):
                        os.makedirs(doc_img_dir)

                    temp_pdf_path = os.path.join(doc_img_dir, f"~temp_{base_name}.pdf")

                    # 1. 中转 PDF 渲染 (调用 Word 将文档锁定排版为 PDF)
                    doc = word_app.Documents.Open(os.path.abspath(file_path), ReadOnly=True)
                    doc.SaveAs(os.path.abspath(temp_pdf_path), FileFormat=17)
                    doc.Close(SaveChanges=0)
                    doc = None

                    # 2. 高清图像切割
                    self.update_tree_row(item_id, (idx, file_name, f".{img_fmt.upper()}",
                                                   f"⚙ 正在以 {dpi_value}DPI 切割图片..."), "processing")
                    pdf_document = fitz.open(temp_pdf_path)
                    total_pages = len(pdf_document)

                    for page_num in range(total_pages):
                        if self._stop_event.is_set():
                            break
                        self._pause_event.wait()

                        page = pdf_document.load_page(page_num)
                        pix = page.get_pixmap(dpi=dpi_value, alpha=False)

                        img_filename = f"{base_name}_{page_num + 1}.{img_fmt}"
                        img_path = os.path.join(doc_img_dir, img_filename)

                        if os.path.exists(img_path):
                            os.remove(img_path)

                        pix.save(img_path)

                    pdf_document.close()

                    # 3. 清理中转 PDF 痕迹
                    if os.path.exists(temp_pdf_path):
                        try:
                            os.remove(temp_pdf_path)
                        except:
                            pass

                    if not self._stop_event.is_set():
                        self.item_outpath_map[item_id] = doc_img_dir
                        self.update_tree_row(item_id,
                                             (idx, file_name, f".{img_fmt.upper()}", f"✔ 成功 (共 {total_pages} 张)"),
                                             "success")
                        success_count += 1

                except Exception as e:
                    error_msg = str(e)
                    kind = com_error_kind(error_msg)
                    if kind == "stuck":
                        friendly_err = "✖ Word 引擎卡死 (请清理后台 WINWORD.EXE)"
                    elif kind == "crash":
                        friendly_err = "✖ Word 崩溃闪退 (原文件可能损坏或排版异常)"
                    else:
                        friendly_err = f"✖ 渲染异常: {error_msg}"

                    self.update_tree_row(item_id, (idx, file_name, f".{img_fmt.upper()}", friendly_err), "error")
                    error_count += 1
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=0)
                        except:
                            pass
                    # 兜底清理未完成的临时文件
                    if temp_pdf_path and os.path.exists(temp_pdf_path):
                        try:
                            os.remove(temp_pdf_path)
                        except:
                            pass

            if self._stop_event.is_set():
                for item_id in pending_items:
                    if self.tree.exists(item_id):
                        tags = self.tree.item(item_id, "tags")
                        if "pending" in tags or "processing" in tags:
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
            kind = com_error_kind(error_msg)
            if kind == "stuck":
                user_msg = "底层 Word 引擎启动失败！请按下 Ctrl+Shift+Esc 结束所有 WINWORD.EXE 进程后重试。"
            else:
                user_msg = f"发生全局错误: {error_msg}\n请确保已安装正版 Microsoft Office Word。"

            self.parent.after(0, lambda e=user_msg: self.set_status("✖", e, "#DC3545"))
            self.parent.after(0, lambda: self.update_button_ui("idle"))

        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
