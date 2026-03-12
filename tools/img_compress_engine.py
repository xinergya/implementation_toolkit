import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, UnidentifiedImageError

# 核心拖拽支持库
from tkinterdnd2 import DND_FILES, TkinterDnD


class ImageCompressUI:
    """图片智能压缩引擎 (支持混合拖拽 + 安全覆盖算法 + 智能格式分流)"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.out_mode_var = tk.StringVar(value="new_dir")
        self.target_var = tk.StringVar()
        self.keep_format_var = tk.BooleanVar(value=True)

        self.item_filepath_map = {}
        self.item_outpath_map = {}

        self.is_running = False
        self.anim_job = None
        self.anim_idx = 0

        self.setup_ui()

        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.on_drag_drop)

    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)  # 让第 5 行的表格区完美拉伸

        # 1. 源图片添加区
        tk.Label(frame, text="1. 添加源图片:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0,
                                                                                             sticky=tk.W, pady=5)
        btn_source_frame = tk.Frame(frame)
        btn_source_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)

        tk.Button(btn_source_frame, text="📄 添加图片", font=("Microsoft YaHei UI", 9),
                  command=self.add_files).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="📁 添加文件夹 (扫描子目录)", font=("Microsoft YaHei UI", 9),
                  command=self.add_folder).pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(btn_source_frame, text="▤ 清空列表", font=("Microsoft YaHei UI", 9),
                  command=self.clear_list).pack(side=tk.LEFT)

        tk.Label(btn_source_frame, text="💡", font=("Microsoft YaHei UI", 12), fg="#F39C12").pack(side=tk.LEFT,
                                                                                                 padx=(15, 0))
        tk.Label(btn_source_frame, text="提示：支持将图片和文件夹直接混搭拖拽到下方列表中。",
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)

        # 2. 【新增】：输出模式选择区
        tk.Label(frame, text="2. 输出模式:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W,
                                                                                           pady=5)
        out_mode_frame = tk.Frame(frame)
        out_mode_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=5)

        tk.Radiobutton(out_mode_frame, text="另存为新文件 (追加 _min 后缀，绝对安全)", variable=self.out_mode_var,
                       value="new_dir", font=("Microsoft YaHei UI", 10), command=self.toggle_output_state).pack(
            side=tk.LEFT, padx=(0, 20))
        tk.Radiobutton(out_mode_frame, text="直接覆盖原文件 (原位瘦身，极度精简)", variable=self.out_mode_var,
                       value="overwrite", font=("Microsoft YaHei UI", 10), fg="#D35400",
                       command=self.toggle_output_state).pack(side=tk.LEFT)

        # 3. 输出目录选择 (UI 联动挂载点)
        self.lbl_target = tk.Label(frame, text="3. 保存目录:", font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_target.grid(row=2, column=0, sticky=tk.W, pady=5)
        self.entry_target = tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10))
        self.entry_target.grid(row=2, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        self.btn_target = tk.Button(frame, text="浏览目录...", command=self.select_target, width=12)
        self.btn_target.grid(row=2, column=2, pady=5)

        # 4. 【升级】：压缩配置区 (新增格式保留开关)
        tk.Label(frame, text="4. 压缩配置:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=3, column=0, sticky=tk.W,
                                                                                           pady=5)
        config_frame = tk.Frame(frame)
        config_frame.grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=10, pady=5)

        self.quality_var = tk.StringVar(value="85 (视觉无损，推荐，体积减小约 40%)")
        self.quality_combo = ttk.Combobox(config_frame, textvariable=self.quality_var, font=("Microsoft YaHei UI", 10),
                                          width=38)
        self.quality_combo['values'] = (
            "95 (轻微压缩，保留极高画质)",
            "85 (视觉无损，推荐，体积减小约 40%)",
            "70 (平衡模式，适合大部分网络传输)",
            "50 (极限压缩，画质略损，体积减小约 70%+)",
            "100 (仅清理冗余数据，几乎不改变画质)"
        )
        self.quality_combo.pack(side=tk.LEFT, ipady=3)

        # 新增的格式保护开关
        tk.Checkbutton(config_frame, text="智能保留原图格式与透明背景 (取消勾选则极限强转 JPG)",
                       variable=self.keep_format_var, font=("Microsoft YaHei UI", 9, "bold"), fg="#2980B9").pack(
            side=tk.LEFT, padx=(15, 0))

        # 5. 控制按钮区
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)

        self.btn_pause = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 6. 可视化数据表格区
        list_frame = tk.Frame(frame)
        list_frame.grid(row=5, column=0, columnspan=3, sticky=tk.NSEW, pady=5)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        columns = ("id", "filename", "orig_size", "new_size", "ratio", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)

        self.tree.heading("id", text="序号")
        self.tree.column("id", width=50, anchor="center")
        self.tree.heading("filename", text="图片名称")
        self.tree.column("filename", width=250, anchor="w")
        self.tree.heading("orig_size", text="原大小")
        self.tree.column("orig_size", width=90, anchor="center")
        self.tree.heading("new_size", text="压缩后")
        self.tree.column("new_size", width=90, anchor="center")
        self.tree.heading("ratio", text="压缩率")
        self.tree.column("ratio", width=90, anchor="center")
        self.tree.heading("status", text="处理状态")
        self.tree.column("status", width=180, anchor="center")

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)

        self.tree.tag_configure("success", foreground="#198754")
        self.tree.tag_configure("error", foreground="#DC3545")
        self.tree.tag_configure("processing", foreground="#0078D7")
        self.tree.tag_configure("pending", foreground="#6C757D")
        self.tree.tag_configure("skipped", foreground="#0D6EFD")

        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Delete>", self.delete_selected_items)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.context_menu = tk.Menu(self.tree, tearoff=0, font=("Microsoft YaHei UI", 10))
        self.context_menu.add_command(label="✖ 移除选中图片 (Del)", command=self.delete_selected_items)

        # 7. 底部汇总栏
        bottom_frame = tk.Frame(frame)
        bottom_frame.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=(10, 0))

        self.lbl_icon_var = tk.StringVar(value="💡")
        self.lbl_icon = tk.Label(bottom_frame, textvariable=self.lbl_icon_var, font=("Microsoft YaHei UI", 12))
        self.lbl_icon.pack(side=tk.LEFT)

        self.lbl_summary_var = tk.StringVar(value="提示：双击表格记录可打开图片；按 Delete 键可踢除图片。")
        self.lbl_summary = tk.Label(bottom_frame, textvariable=self.lbl_summary_var,
                                    font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_summary.pack(side=tk.LEFT)

        self.btn_open_dir = tk.Button(bottom_frame, text="📂 打开输出目录", font=("Microsoft YaHei UI", 9),
                                      command=self.open_output_dir)
        self.btn_open_dir.pack(side=tk.RIGHT)

        self.update_button_ui("idle")
        self.toggle_output_state()  # 初始化 UI 状态
        self.set_status("💡", "提示：双击表格记录可打开图片；按 Delete 键可踢除图片。", "#F39C12", "#2C3E50")

    # ================= 状态与 UI 调度引擎 =================
    def toggle_output_state(self):
        """当选择覆盖模式时，智能置灰保存目录相关控件"""
        if self.out_mode_var.get() == "overwrite":
            self.entry_target.config(state=tk.DISABLED)
            self.btn_target.config(state=tk.DISABLED)
            self.lbl_target.config(fg="#ADB5BD")
            self.btn_open_dir.config(state=tk.DISABLED)
        else:
            self.entry_target.config(state=tk.NORMAL)
            self.btn_target.config(state=tk.NORMAL)
            self.lbl_target.config(fg="black")
            if not self.is_running:
                self.btn_open_dir.config(state=tk.NORMAL)

    def set_status(self, icon, text, icon_color, text_color=None):
        self.lbl_icon_var.set(icon)
        self.lbl_icon.config(fg=icon_color)
        self.lbl_summary_var.set(text)
        self.lbl_summary.config(fg=text_color if text_color else icon_color)

    # ================= 动态添加文件与数据源引擎 =================
    def is_valid_image_file(self, filename):
        return filename.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.webp'))

    def on_drag_drop(self, event):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "任务正在运行，请先暂停后再拖入新图片！")
            return

        dropped_paths = self.tree.tk.splitlist(event.data)
        images_to_add = []

        for path in dropped_paths:
            if os.path.isfile(path) and self.is_valid_image_file(path):
                images_to_add.append(path)
            elif os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for f in files:
                        if self.is_valid_image_file(f):
                            images_to_add.append(os.path.join(root, f))

        if images_to_add:
            self._append_to_queue(images_to_add)
            if not self.target_var.get():
                self.target_var.set(os.path.join(os.path.dirname(images_to_add[0]), "压缩后输出"))
        else:
            messagebox.showinfo("提示", "拖入的项目中没有找到支持的图片文件！")

    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择图片文件",
                                            filetypes=[("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.webp")])
        if paths:
            valid_paths = [p for p in paths if self.is_valid_image_file(p)]
            self._append_to_queue(valid_paths)
            if valid_paths and not self.target_var.get():
                self.target_var.set(os.path.join(os.path.dirname(valid_paths[0]), "压缩后输出"))

    def add_folder(self):
        path = filedialog.askdirectory(title="选择包含图片的文件夹")
        if path:
            found_images = []
            for root, dirs, files in os.walk(path):
                for f in files:
                    if self.is_valid_image_file(f):
                        found_images.append(os.path.join(root, f))
            if found_images:
                self._append_to_queue(found_images)
                if not self.target_var.get():
                    self.target_var.set(os.path.join(path, "压缩后输出"))
            else:
                messagebox.showinfo("扫描结果", "该目录下未找到支持的图片文件！")

    def _append_to_queue(self, paths):
        existing_paths = list(self.item_filepath_map.values())
        added_count = 0

        for path in paths:
            if path not in existing_paths:
                idx = len(self.tree.get_children()) + 1
                file_name = os.path.basename(path)
                try:
                    orig_size = os.path.getsize(path)
                    size_str = self.format_size(orig_size)
                except Exception:
                    size_str = "读取失败"

                item_id = self.tree.insert("", tk.END, values=(idx, file_name, size_str, "-", "-", "⌛ 等待中"),
                                           tags=("pending",))
                self.item_filepath_map[item_id] = path
                added_count += 1

        self.set_status("✔", f"成功加载 {added_count} 张图片。队列中共 {len(self.tree.get_children())} 张。", "#198754")

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
            frames = ["⚙ 智能压缩中.  ", "⚙ 智能压缩中.. ", "⚙ 智能压缩中..."]
            self.btn_start.config(text=frames[self.anim_idx % len(frames)])
            self.anim_idx += 1
            self.anim_job = self.parent.after(400, self.start_btn_animation)

    def stop_btn_animation(self):
        if self.anim_job is not None:
            self.parent.after_cancel(self.anim_job)
            self.anim_job = None

    def update_button_ui(self, state):
        color_start_active = "#1ABC9C"
        color_pause_active = "#F39C12"
        color_stop_active = "#DC3545"
        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray_bg = "#E0E0E0"
        color_enabled_gray_fg = "#333333"

        self.stop_btn_animation()

        if state == "idle":
            self.btn_start.config(bg=color_start_active, fg="white", text="▶ 开始智能压缩", state=tk.NORMAL,
                                  command=self.start_compression_thread)
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
            self.btn_pause.config(bg=color_pause_active, fg="white", text="▶ 继续压缩", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="■ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)

        elif state == "stopping":
            self.btn_start.config(bg=color_disabled_bg, fg=color_disabled_fg, text="■ 正在中止...", state=tk.DISABLED)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="❚❚ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="■ 正在停止...", state=tk.NORMAL,
                                   command=lambda: None)

    # ================= 辅助交互响应 =================
    def show_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def delete_selected_items(self, event=None):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "任务正在运行中，请先暂停后再执行删除操作！")
            return
        selected_items = self.tree.selection()
        if not selected_items: return

        for item_id in selected_items:
            self.tree.delete(item_id)
            if item_id in self.item_filepath_map: del self.item_filepath_map[item_id]
            if item_id in self.item_outpath_map: del self.item_outpath_map[item_id]

        for idx, item_id in enumerate(self.tree.get_children(), 1):
            vals = list(self.tree.item(item_id, "values"))
            vals[0] = idx
            self.tree.item(item_id, values=vals)

        self.set_status("✖", f"已剔除选中图片。当前列表剩余 {len(self.tree.get_children())} 张。", "#6C757D", "#2C3E50")

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
                messagebox.showerror("错误", f"无法打开文件:\n{str(e)}")

    def select_target(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path: self.target_var.set(path)

    def open_output_dir(self):
        tgt = self.target_var.get().strip()
        if os.path.exists(tgt): os.startfile(tgt)

    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.update_button_ui("paused")
            self.set_status("❚❚", "任务已挂起，您可以随时删除待办列表中的图片...", "#D35400")
        else:
            self._pause_event.set()
            self.update_button_ui("running")
            self.set_status("▶", "任务已恢复，正在全力压缩...", "#0078D7")

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", "确定要中止当前的压缩任务吗？\n已压缩完成的图片将被保留。"):
            self._stop_event.set()
            self._pause_event.set()
            self.update_button_ui("stopping")
            self.set_status("■", "正在安全停止任务...", "#DC3545")

    def format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.2f} MB"

    def update_tree_row(self, item_id, values, tag, auto_scroll=True):
        def update():
            if self.tree.exists(item_id):
                self.tree.item(item_id, values=values, tags=(tag,))
                if auto_scroll: self.tree.see(item_id)

        self.parent.after(0, update)

    def start_compression_thread(self):
        out_mode = self.out_mode_var.get()
        tgt = self.target_var.get().strip().strip('"').strip("'")
        keep_format = self.keep_format_var.get()

        raw_q = self.quality_var.get().strip()
        q_str = raw_q.split(' ')[0]
        try:
            q_value = int(q_str)
            if not (1 <= q_value <= 100): raise ValueError
        except ValueError:
            messagebox.showwarning("参数错误", "请输入有效的压缩质量数值(1-100)！")
            return

        if out_mode == "new_dir" and not tgt:
            messagebox.showwarning("提示", "您选择了另存为新目录，请先选择输出保存位置！")
            return

        pending_items = list(self.tree.get_children())
        if not pending_items:
            messagebox.showinfo("提示", "列表中没有待处理的图片！")
            return

        if out_mode == "new_dir" and os.path.exists(tgt) and os.listdir(tgt):
            if not messagebox.askyesno("防覆盖提示",
                                       "⚠️ 输出目录非空！\n\n继续压缩可能会覆盖里面已存在的同名图片。\n是否确认继续执行？"):
                return
        elif out_mode == "overwrite":
            if not messagebox.askyesno("高危操作确认",
                                       "⚠️ 您选择了【直接覆盖原文件】！\n\n压缩成功后，原图片将被彻底替换且无法通过回收站找回。\n是否确认执行覆盖操作？"):
                return

        self.item_outpath_map.clear()
        for item_id in pending_items:
            vals = list(self.tree.item(item_id, "values"))
            vals[3] = "-"
            vals[4] = "-"
            vals[5] = "⌛ 等待中"
            self.tree.item(item_id, values=vals, tags=("pending",))

        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        self.update_button_ui("running")
        if out_mode == "new_dir": self.btn_open_dir.config(state=tk.DISABLED)

        self.set_status("⚙", f"开始执行，队列中共有 {len(pending_items)} 张图片...", "#0078D7")

        threading.Thread(target=self.process_compression, args=(pending_items, tgt, q_value, out_mode, keep_format),
                         daemon=True).start()

    # ================= 核心压缩与 【工业级安全覆盖算法】 =================
    def process_compression(self, pending_items, tgt, q_value, out_mode, keep_format):
        success_count = 0
        error_count = 0
        total_saved_bytes = 0

        try:
            for item_id in pending_items:
                if self._stop_event.is_set(): break
                self._pause_event.wait()

                if not self.tree.exists(item_id) or item_id not in self.item_filepath_map: continue

                file_path = self.item_filepath_map[item_id]
                file_name = os.path.basename(file_path)
                idx = self.tree.item(item_id, "values")[0]

                try:
                    orig_size = os.path.getsize(file_path)
                except Exception:
                    self.update_tree_row(item_id, (idx, file_name, "读取失败", "-", "-", "✖ 文件不存在或被占用"),
                                         "error")
                    error_count += 1
                    continue

                self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size), "-", "-", "⚙ 处理中..."),
                                     "processing")

                temp_path = ""
                try:
                    orig_ext = os.path.splitext(file_name)[1].lower()
                    base_name = os.path.splitext(file_name)[0]

                    # 决定工作目录
                    if out_mode == "new_dir":
                        current_out_dir = tgt
                        if not os.path.exists(current_out_dir): os.makedirs(current_out_dir)
                    else:
                        current_out_dir = os.path.dirname(file_path)  # 原位覆盖模式

                    # 1. 创建基于时间戳的绝对安全临时文件
                    temp_filename = f"~temp_{base_name}_{int(time.time() * 100)}.tmp"
                    temp_path = os.path.join(current_out_dir, temp_filename)

                    # 2. 图像智能分流渲染
                    with Image.open(file_path) as img:
                        save_format = img.format if img.format else "JPEG"

                        if not keep_format:
                            # 强转极限压榨模式
                            if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                            save_format = "JPEG"
                            out_ext = ".jpg"
                        else:
                            # 智能保留模式
                            out_ext = orig_ext
                            if save_format == "JPEG" and img.mode in ("RGBA", "P"):
                                img = img.convert("RGB")  # 纠正错误的标头

                        # 执行物理存储至临时文件
                        if save_format == "PNG":
                            img.save(temp_path, format=save_format, optimize=True)
                        else:
                            try:
                                img.save(temp_path, format=save_format, quality=q_value, optimize=True)
                            except:
                                img.save(temp_path, format=save_format, optimize=True)

                    new_size = os.path.getsize(temp_path)

                    # 3. 智判定：体积是否真的变小了？
                    if new_size < orig_size:
                        saved_bytes = orig_size - new_size
                        ratio = (saved_bytes / orig_size) * 100
                        total_saved_bytes += saved_bytes

                        if out_mode == "overwrite":
                            final_path = os.path.join(current_out_dir, base_name + out_ext)
                            # 如果强转了格式(png->jpg)，需手动删掉原 png 文件
                            if os.path.abspath(file_path) != os.path.abspath(final_path):
                                os.remove(file_path)
                            elif os.path.exists(final_path):
                                os.remove(final_path)  # 删除同名原文件

                            os.rename(temp_path, final_path)  # 偷梁换柱
                            self.item_outpath_map[item_id] = final_path
                            self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size),
                                                           self.format_size(new_size), f"↓ {ratio:.1f}%",
                                                           "✔ 原位覆盖成功"), "success")
                        else:
                            final_path = os.path.join(current_out_dir, f"{base_name}_min{out_ext}")
                            if os.path.exists(final_path): os.remove(final_path)
                            os.rename(temp_path, final_path)
                            self.item_outpath_map[item_id] = final_path
                            self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size),
                                                           self.format_size(new_size), f"↓ {ratio:.1f}%",
                                                           "✔ 另存压缩成功"), "success")

                        temp_path = ""  # 成功后清空路径，防止 finally 误删
                        success_count += 1
                    else:
                        # 如果压缩后反而变大，抛弃临时文件，执行智能跳过
                        os.remove(temp_path)
                        temp_path = ""

                        if out_mode == "overwrite":
                            self.item_outpath_map[item_id] = file_path
                            self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size),
                                                           self.format_size(orig_size), "0.0%", "✨ 跳过 (原图最优)"),
                                                 "skipped")
                        else:
                            final_path = os.path.join(current_out_dir, f"{base_name}_min{orig_ext}")
                            shutil.copy2(file_path, final_path)
                            self.item_outpath_map[item_id] = final_path
                            self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size),
                                                           self.format_size(orig_size), "0.0%", "✨ 复制 (原图最优)"),
                                                 "skipped")

                        success_count += 1

                except UnidentifiedImageError:
                    self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size), "N/A", "N/A",
                                                   "✖ 图片损坏或伪装"), "error")
                    error_count += 1
                except PermissionError:
                    self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size), "N/A", "N/A",
                                                   "✖ 无权限或被占用"), "error")
                    error_count += 1
                except Exception as e:
                    self.update_tree_row(item_id, (idx, file_name, self.format_size(orig_size), "N/A", "N/A",
                                                   f"✖ 异常: {str(e)}"), "error")
                    error_count += 1
                finally:
                    # 绝对防漏清理机制：无论发生什么报错，保证残留的隐藏临时文件被销毁
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except:
                            pass

            if self._stop_event.is_set():
                for item_id in pending_items:
                    if self.tree.exists(item_id):
                        tags = self.tree.item(item_id, "tags")
                        if "pending" in tags:
                            vals = list(self.tree.item(item_id, "values"))
                            vals[5] = "✖ 已取消"
                            self.update_tree_row(item_id, vals, "error", auto_scroll=False)

            def finish_ui():
                self.is_running = False
                self.update_button_ui("idle")
                self.toggle_output_state()  # 还原目录按钮状态

                if self._stop_event.is_set():
                    self.set_status("■", "任务已中断。", "#DC3545")
                else:
                    self.set_status("✔",
                                    f"处理完成！成功 {success_count} 张，失败 {error_count} 张。共节省 {self.format_size(total_saved_bytes)}！",
                                    "#198754")

            self.parent.after(0, finish_ui)

        except Exception as global_e:
            self.parent.after(0, lambda e=str(global_e): self.set_status("✖", f"发生全局严重错误: {e}", "#DC3545"))
            self.parent.after(0, lambda: self.update_button_ui("idle"))
            self.parent.after(0, self.toggle_output_state)


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("图片智能压缩 极速引擎")
    root.geometry("900x650")
    app = ImageCompressUI(root)
    root.mainloop()