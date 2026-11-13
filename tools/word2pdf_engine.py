import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 核心：操作 Windows 系统的底层库
import win32com.client
import pythoncom
# 核心拖拽支持库
from tkinterdnd2 import DND_FILES, TkinterDnD

from utils.tree_ui_base import TreeEngineUI
from utils.constants import com_error_kind
from utils.pathutil import is_valid_word_file


class WordToPdfUI(TreeEngineUI):
    """Word 转 PDF 处理引擎 (支持混合拖拽 + 动态反馈 + 优雅排版保护)"""

    COLOR_START = "#1ABC9C"
    START_TEXT = "▶ 开始转换 PDF"
    RESUME_TEXT = "▶ 继续转换"
    ANIM_FRAMES = ("⚙ Word 引擎运作中.  ", "⚙ Word 引擎运作中.. ", "⚙ Word 引擎运作中...")
    CANCEL_CONFIRM = "确定要中止当前的转换任务吗？\n已转换完成的 PDF 将被保留。"
    STOPPING_STATUS_TEXT = "正在安全停止任务，请等待后台 Word 引擎释放资源..."
    PAUSE_STATUS_TEXT = "任务已挂起，您可以随时删除待办列表中的文档..."
    RESUME_STATUS_TEXT = "任务已恢复，正在全力转换 PDF..."
    DRAG_RUNNING_WARN = "任务正在运行，请先暂停后再拖入新文档！"
    DROP_EMPTY_MSG = "拖入的项目中没有找到有效的 Word 文档！"
    ADD_FILES_TITLE = "选择 Word 文档"
    ADD_FILES_FILETYPES = [("Word 文档", "*.doc;*.docx")]
    ADD_FOLDER_TITLE = "选择包含 Word 的文件夹"
    ADD_FOLDER_EMPTY_MSG = "该目录下未找到有效的 Word 文件！"

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
        frame.rowconfigure(3, weight=1)  # 让表格区完美拉伸

        # 1. 源文档添加区 (合并文件与文件夹功能)
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
        tk.Label(btn_source_frame, text="提示：支持将 Word 文件和文件夹直接混搭拖拽到下方列表中。",
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        # 2. 输出目录选择 (精简布局)
        tk.Label(frame, text="2. PDF 保存至:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10)).grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.select_target, width=12).grid(row=1, column=2, pady=5)

        # 3. 控制按钮区 (引入状态机管理)
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

        self.lbl_summary_var = tk.StringVar(value="提示：双击表格记录可打开 PDF；按 Delete 键可踢除文档。")
        self.lbl_summary = tk.Label(bottom_frame, textvariable=self.lbl_summary_var, font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_summary.pack(side=tk.LEFT)

        self.btn_open_dir = tk.Button(bottom_frame, text="📂 打开输出目录", font=("Microsoft YaHei UI", 9),
                                      command=self.open_output_dir, state=tk.DISABLED)
        self.btn_open_dir.pack(side=tk.RIGHT)

        self.update_button_ui("idle")
        self.set_status("💡", "提示：双击表格记录可打开 PDF；按 Delete 键可踢除文档。", "#F39C12", "#2C3E50")

    def is_valid_file(self, path):
        return is_valid_word_file(path)

    def build_row_values(self, idx, path):
        return (idx, os.path.basename(path), ".PDF", "⌛ 等待中")

    def start_conversion_thread(self):
        tgt = self.target_var.get().strip().strip('"').strip("'")

        if not tgt:
            messagebox.showwarning("提示", "请先选择输出保存位置！")
            return

        pending_items = list(self.tree.get_children())
        if not pending_items:
            messagebox.showinfo("提示", "列表中没有待处理的文档！可以点击上方添加或直接拖入文件。")
            return

        # 防覆盖检查
        if os.path.exists(tgt) and os.listdir(tgt):
            if not messagebox.askyesno("防覆盖提示", "⚠️ 输出目录非空！\n\n继续转换可能会直接覆盖里面已存在的同名 PDF。\n是否确认覆盖并继续执行？"):
                return

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

        self.set_status("⚙", f"正在唤醒 Word 引擎，队列中共 {len(pending_items)} 份文档...", "#0078D7")

        threading.Thread(target=self.process_conversion, args=(pending_items, tgt), daemon=True).start()

    # ================= 核心转换与优雅覆盖逻辑 =================
    def process_conversion(self, pending_items, tgt):
        pythoncom.CoInitialize()
        word_app = None
        success_count = 0
        error_count = 0

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

                self.update_tree_row(item_id, (idx, file_name, ".PDF", "🔄 正在调用 Word 转换..."), "processing")

                doc = None
                try:
                    if not os.path.exists(tgt):
                        os.makedirs(tgt)

                    base_name = os.path.splitext(file_name)[0]
                    pdf_path = os.path.join(tgt, f"{base_name}.pdf")

                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)

                    doc = word_app.Documents.Open(file_path, ReadOnly=True)
                    # FileFormat=17 代表 wdFormatPDF
                    doc.SaveAs(pdf_path, FileFormat=17)

                    self.item_outpath_map[item_id] = pdf_path
                    self.update_tree_row(item_id, (idx, file_name, ".PDF", "✔ 转换成功"), "success")
                    success_count += 1

                except PermissionError:
                    self.update_tree_row(item_id, (idx, file_name, ".PDF", "✖ 覆盖失败: 请先关闭正打开该PDF的软件"), "error")
                    error_count += 1

                except Exception as e:
                    error_msg = str(e)
                    kind = com_error_kind(error_msg)
                    if kind == "stuck":
                        friendly_err = "✖ 引擎卡死 (请在任务管理器清理 WINWORD.EXE)"
                    elif kind == "crash":
                        friendly_err = "✖ Word 崩溃闪退 (文档可能损坏或过大)"
                    else:
                        friendly_err = "✖ 格式损坏或含无法转换的内容"

                    self.update_tree_row(item_id, (idx, file_name, ".PDF", friendly_err), "error")
                    error_count += 1

                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=0)
                        except:
                            pass

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
            kind = com_error_kind(error_msg)
            if kind == "stuck":
                user_msg = "底层 Word 引擎启动失败！请按下 Ctrl+Shift+Esc 结束所有 WINWORD.EXE 进程后重试。"
            else:
                user_msg = f"发生全局严重错误: {error_msg}\n请确保已安装正版 Microsoft Office Word。"

            self.parent.after(0, lambda e=user_msg: self.set_status("✖", e, "#DC3545"))
            self.parent.after(0, lambda: self.update_button_ui("idle"))

        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
