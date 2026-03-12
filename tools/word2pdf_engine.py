import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 核心：操作 Windows 系统的底层库
import win32com.client
import pythoncom
# 核心拖拽支持库
from tkinterdnd2 import DND_FILES, TkinterDnD


class WordToPdfUI:
    """Word 转 PDF 处理引擎 (支持混合拖拽 + 动态反馈 + 优雅排版保护)"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.target_var = tk.StringVar()

        # 映射架构
        self.item_filepath_map = {}  # 表格行ID -> 原始Word路径
        self.item_outpath_map = {}   # 表格行ID -> 生成的PDF路径

        self.is_running = False
        self.anim_job = None  # 动态按钮动画任务
        self.anim_idx = 0     # 动态按钮帧索引

        self.setup_ui()

        # 【核心功能】：注册拖拽事件支持
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

        # 【完美拆分】：专门渲染顶部提示图标的 Label (橘黄色生效)
        tk.Label(btn_source_frame, text="💡", font=("Microsoft YaHei UI", 12), fg="#F39C12").pack(side=tk.LEFT, padx=(15, 0))
        # 专门渲染后续提示文字的 Label
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

        # 5. 底部汇总栏 (深度改造区)
        bottom_frame = tk.Frame(frame)
        bottom_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(10, 0))

        # 拆分 1：专属图标 Label (独立控制颜色)
        self.lbl_icon_var = tk.StringVar(value="💡")
        self.lbl_icon = tk.Label(bottom_frame, textvariable=self.lbl_icon_var, font=("Microsoft YaHei UI", 12))
        self.lbl_icon.pack(side=tk.LEFT)

        # 拆分 2：专属文本 Label
        self.lbl_summary_var = tk.StringVar(value="提示：双击表格记录可打开 PDF；按 Delete 键可踢除文档。")
        self.lbl_summary = tk.Label(bottom_frame, textvariable=self.lbl_summary_var, font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_summary.pack(side=tk.LEFT)

        self.btn_open_dir = tk.Button(bottom_frame, text="📂 打开输出目录", font=("Microsoft YaHei UI", 9),
                                      command=self.open_output_dir, state=tk.DISABLED)
        self.btn_open_dir.pack(side=tk.RIGHT)

        self.update_button_ui("idle")
        self.set_status("💡", "提示：双击表格记录可打开 PDF；按 Delete 键可踢除文档。", "#F39C12", "#2C3E50")

    # ================= 新增：底部状态栏调度中心 =================
    def set_status(self, icon, text, icon_color, text_color=None):
        """统一管理底部状态栏的图标与文字分离渲染"""
        self.lbl_icon_var.set(icon)
        self.lbl_icon.config(fg=icon_color)

        self.lbl_summary_var.set(text)
        if text_color:
            self.lbl_summary.config(fg=text_color)
        else:
            self.lbl_summary.config(fg=icon_color)

    # ================= 动态添加文件与数据源引擎 =================
    def is_valid_word_file(self, filename):
        """校验是否为合法的 Word 文档且非隐藏临时文件"""
        basename = os.path.basename(filename)
        return filename.lower().endswith(('.doc', '.docx')) and not basename.startswith('~$')

    def on_drag_drop(self, event):
        """处理拖拽进来的混合数据（文件 + 文件夹同时处理）"""
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "任务正在运行，请先暂停后再拖入新文档！")
            return

        dropped_paths = self.tree.tk.splitlist(event.data)
        words_to_add = []

        for path in dropped_paths:
            if os.path.isfile(path) and self.is_valid_word_file(path):
                words_to_add.append(path)
            elif os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for f in files:
                        if self.is_valid_word_file(f):
                            words_to_add.append(os.path.join(root, f))

        if words_to_add:
            self._append_to_queue(words_to_add)
            if not self.target_var.get():
                self.target_var.set(os.path.dirname(words_to_add[0]))
        else:
            messagebox.showinfo("提示", "拖入的项目中没有找到有效的 Word 文档！")

    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择 Word 文档", filetypes=[("Word 文档", "*.doc;*.docx")])
        if paths:
            valid_paths = [p for p in paths if self.is_valid_word_file(p)]
            self._append_to_queue(valid_paths)
            if valid_paths and not self.target_var.get():
                self.target_var.set(os.path.dirname(valid_paths[0]))

    def add_folder(self):
        path = filedialog.askdirectory(title="选择包含 Word 的文件夹")
        if path:
            found_words = []
            for root, dirs, files in os.walk(path):
                for f in files:
                    if self.is_valid_word_file(f):
                        found_words.append(os.path.join(root, f))
            if found_words:
                self._append_to_queue(found_words)
                if not self.target_var.get():
                    self.target_var.set(path)
            else:
                messagebox.showinfo("扫描结果", "该目录下未找到有效的 Word 文件！")

    def _append_to_queue(self, paths):
        existing_paths = list(self.item_filepath_map.values())
        added_count = 0

        for path in paths:
            if path not in existing_paths:
                idx = len(self.tree.get_children()) + 1
                file_name = os.path.basename(path)
                item_id = self.tree.insert("", tk.END, values=(idx, file_name, ".PDF", "⌛ 等待中"),
                                           tags=("pending",))
                self.item_filepath_map[item_id] = path
                added_count += 1

        self.set_status("✔", f"成功添加 {added_count} 份文档。队列中共 {len(self.tree.get_children())} 份。", "#198754")

    def clear_list(self):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "任务正在运行，请先暂停！")
            return
        self.tree.delete(*self.tree.get_children())
        self.item_filepath_map.clear()
        self.item_outpath_map.clear()
        self.set_status("▤", "列表已清空。您可以点击按钮或直接拖拽文件进来。", "#6C757D", "#2C3E50")

    # ================= 动态按钮与UI状态机 =================
    def start_btn_animation(self):
        if self.is_running and self._pause_event.is_set():
            frames = ["⚙ Word 引擎运作中.  ", "⚙ Word 引擎运作中.. ", "⚙ Word 引擎运作中..."]
            self.btn_start.config(text=frames[self.anim_idx % len(frames)])
            self.anim_idx += 1
            self.anim_job = self.parent.after(400, self.start_btn_animation)

    def stop_btn_animation(self):
        if self.anim_job is not None:
            self.parent.after_cancel(self.anim_job)
            self.anim_job = None

    def update_button_ui(self, state):
        color_start_active = "#1ABC9C"  # 绿色 (Word 转 PDF 专用色)
        color_pause_active = "#F39C12"
        color_stop_active = "#DC3545"
        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray_bg = "#E0E0E0"
        color_enabled_gray_fg = "#333333"

        self.stop_btn_animation()

        if state == "idle":
            self.btn_start.config(bg=color_start_active, fg="white", text="▶ 开始转换 PDF", state=tk.NORMAL,
                                  command=self.start_conversion_thread)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="❚❚ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_disabled_bg, fg=color_disabled_fg, text="■ 停止", state=tk.DISABLED)

        elif state == "running":
            self.btn_start.config(bg=color_start_active, fg="white", state=tk.NORMAL, command=lambda: None)
            self.btn_pause.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="❚❚ 暂停", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="■ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)
            self.start_btn_animation()

        elif state == "paused":
            self.btn_start.config(bg=color_enabled_gray_bg, fg="#888888", text="⚙ 已挂起", state=tk.NORMAL,
                                  command=lambda: None)
            self.btn_pause.config(bg=color_pause_active, fg="white", text="▶ 继续转换", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="■ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)

        elif state == "stopping":
            self.btn_start.config(bg=color_disabled_bg, fg=color_disabled_fg, text="■ 正在中止...", state=tk.DISABLED)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="❚❚ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="■ 正在停止...", state=tk.NORMAL,
                                   command=lambda: None)

    # ================= UI 交互响应 =================
    def show_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def delete_selected_items(self, event=None):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "转换任务正在运行中，请先暂停后再执行删除操作！")
            return
        selected_items = self.tree.selection()
        if not selected_items: return

        for item_id in selected_items:
            self.tree.delete(item_id)
            if item_id in self.item_filepath_map:
                del self.item_filepath_map[item_id]
            if item_id in self.item_outpath_map:
                del self.item_outpath_map[item_id]

        for idx, item_id in enumerate(self.tree.get_children(), 1):
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = idx
            self.tree.item(item_id, values=vals)

        self.set_status("✖", f"已剔除选中文档。当前列表剩余 {len(self.tree.get_children())} 份。", "#6C757D", "#2C3E50")

    def on_tree_double_click(self, event):
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        out_path = self.item_outpath_map.get(item_id)

        if not out_path or not os.path.exists(out_path):
            out_path = self.item_filepath_map.get(item_id)

        if out_path and os.path.exists(out_path):
            try:
                os.startfile(out_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文档:\n{str(e)}")

    def select_target(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path: self.target_var.set(path)

    def open_output_dir(self):
        tgt = self.target_var.get().strip()
        if os.path.exists(tgt):
            os.startfile(tgt)

    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.update_button_ui("paused")
            self.set_status("❚❚", "任务已挂起，您可以随时删除待办列表中的文档...", "#D35400")
        else:
            self._pause_event.set()
            self.update_button_ui("running")
            self.set_status("▶", "任务已恢复，正在全力转换 PDF...", "#0078D7")

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", "确定要中止当前的转换任务吗？\n已转换完成的 PDF 将被保留。"):
            self._stop_event.set()
            self._pause_event.set()
            self.update_button_ui("stopping")
            self.set_status("■", "正在安全停止任务，请等待后台 Word 引擎释放资源...", "#DC3545")

    def update_tree_row(self, item_id, values, tag, auto_scroll=True):
        def update():
            if self.tree.exists(item_id):
                self.tree.item(item_id, values=values, tags=(tag,))
                if auto_scroll: self.tree.see(item_id)

        self.parent.after(0, update)

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
                if self._stop_event.is_set(): break
                self._pause_event.wait()

                if not self.tree.exists(item_id) or item_id not in self.item_filepath_map:
                    continue

                file_path = self.item_filepath_map[item_id]
                file_name = os.path.basename(file_path)
                idx = self.tree.item(item_id, "values")[0]

                self.update_tree_row(item_id, (idx, file_name, ".PDF", "🔄 正在调用 Word 转换..."), "processing")

                doc = None
                try:
                    # 混合导入模式下，直接将所有 PDF 输出到选定的目标文件夹
                    if not os.path.exists(tgt):
                        os.makedirs(tgt)

                    base_name = os.path.splitext(file_name)[0]
                    pdf_path = os.path.join(tgt, f"{base_name}.pdf")

                    # 【优雅版：直接移除，依靠原生异常机制防止覆盖卡死】
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)  # 如果文件正被打开，这里会自动触发 PermissionError

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
                    if "80080005" in error_msg or "服务器运行失败" in error_msg:
                        friendly_err = "✖ 引擎卡死 (请在任务管理器清理 WINWORD.EXE)"
                    elif "800706BA" in error_msg or "RPC 服务器不可用" in error_msg:
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
            if "80080005" in error_msg or "服务器运行失败" in error_msg:
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

