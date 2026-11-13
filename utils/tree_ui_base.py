"""树形表格引擎 UI 基类：封装底部状态栏、按钮动画、状态机、树形表格通用操作与拖拽入列。

继承自 BaseEngineUI，适用于「文件列表 + 树形表格 + 暂停/停止状态机」的引擎
（pdf2img / word2img / img_compress / word2pdf）。

子类需在 setup_ui() 中创建以下控件，并实现 is_valid_file / build_row_values：
- self.tree（ttk.Treeview）
- self.btn_start / self.btn_pause / self.btn_cancel
- self.lbl_icon_var / self.lbl_icon（底部状态栏图标）
- self.lbl_summary_var / self.lbl_summary（底部状态栏文字）
- self.target_var、self.btn_open_dir、self.context_menu
并实现 start_conversion_thread()（参数校验后启动后台线程）与 process_conversion(...)。
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from utils.ui_base import BaseEngineUI
from utils.constants import (
    COLOR_DISABLED_BG, COLOR_DISABLED_FG, COLOR_GRAY_BG, COLOR_GRAY_FG,
    COLOR_SUCCESS, COLOR_WARNING, COLOR_MUTED, COLOR_LINK, COLOR_ERROR,
)


class TreeEngineUI(BaseEngineUI):
    # ---------- 子类可覆盖的类属性 ----------
    COLOR_START = COLOR_LINK
    START_TEXT = "▶ 开始"
    RESUME_TEXT = "▶ 继续"
    ANIM_FRAMES = ("⚙ 处理中.  ", "⚙ 处理中.. ", "⚙ 处理中...")
    ITEM_UNIT = "份"        # 队列计数单位（份/张）
    ITEM_LABEL = "文档"     # 项目名称（文档/图片）
    CANCEL_CONFIRM = "确定要中止当前任务吗？\n已完成的结果将被保留。"
    STOPPING_STATUS_TEXT = "正在安全停止任务..."
    PAUSE_STATUS_TEXT = "任务已挂起，您可以随时删除待办列表中的项目..."
    RESUME_STATUS_TEXT = "任务已恢复..."
    DRAG_RUNNING_WARN = "任务正在运行，请先暂停后再拖入新文件！"
    DROP_EMPTY_MSG = "拖入的项目中没有找到有效的文件！"
    ADD_FILES_TITLE = "选择文件"
    ADD_FILES_FILETYPES = [("所有文件", "*.*")]
    ADD_FOLDER_TITLE = "选择文件夹"
    ADD_FOLDER_EMPTY_MSG = "该目录下未找到有效的文件！"

    def __init__(self, parent_frame):
        super().__init__(parent_frame)
        self.is_running = False
        self.anim_job = None
        self.anim_idx = 0
        self.item_filepath_map = {}
        self.item_outpath_map = {}

    # ---------- 底部状态栏 ----------
    def set_status(self, icon, text, icon_color, text_color=None):
        self.lbl_icon_var.set(icon)
        self.lbl_icon.config(fg=icon_color)
        self.lbl_summary_var.set(text)
        self.lbl_summary.config(fg=text_color if text_color else icon_color)

    # ---------- 按钮动画 ----------
    def start_btn_animation(self):
        if self.is_running and self._pause_event.is_set():
            self.btn_start.config(text=self.ANIM_FRAMES[self.anim_idx % len(self.ANIM_FRAMES)])
            self.anim_idx += 1
            self.anim_job = self.parent.after(400, self.start_btn_animation)

    def stop_btn_animation(self):
        if self.anim_job is not None:
            self.parent.after_cancel(self.anim_job)
            self.anim_job = None

    # ---------- 按钮状态机 ----------
    def update_button_ui(self, state):
        self.stop_btn_animation()

        if state == "idle":
            self.btn_start.config(bg=self.COLOR_START, fg="white", text=self.START_TEXT, state=tk.NORMAL,
                                  command=self.start_conversion_thread)
            self.btn_pause.config(bg=COLOR_DISABLED_BG, fg=COLOR_DISABLED_FG, text="❚❚ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=COLOR_DISABLED_BG, fg=COLOR_DISABLED_FG, text="■ 停止", state=tk.DISABLED)
        elif state == "running":
            self.btn_start.config(bg=self.COLOR_START, fg="white", state=tk.NORMAL, command=lambda: None)
            self.btn_pause.config(bg=COLOR_GRAY_BG, fg=COLOR_GRAY_FG, text="❚❚ 暂停", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=COLOR_GRAY_BG, fg=COLOR_GRAY_FG, text="■ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)
            self.start_btn_animation()
        elif state == "paused":
            self.btn_start.config(bg=COLOR_GRAY_BG, fg="#888888", text="⚙ 已挂起", state=tk.NORMAL, command=lambda: None)
            self.btn_pause.config(bg=COLOR_WARNING, fg="white", text=self.RESUME_TEXT, state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=COLOR_GRAY_BG, fg=COLOR_GRAY_FG, text="■ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)
        elif state == "stopping":
            self.btn_start.config(bg=COLOR_DISABLED_BG, fg=COLOR_DISABLED_FG, text="■ 正在中止...", state=tk.DISABLED)
            self.btn_pause.config(bg=COLOR_DISABLED_BG, fg=COLOR_DISABLED_FG, text="❚❚ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=COLOR_ERROR, fg="white", text="■ 正在停止...", state=tk.NORMAL,
                                   command=lambda: None)

    # ---------- 暂停 / 停止 ----------
    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.update_button_ui("paused")
            self.set_status("❚❚", self.PAUSE_STATUS_TEXT, COLOR_WARNING)
        else:
            self._pause_event.set()
            self.update_button_ui("running")
            self.set_status("▶", self.RESUME_STATUS_TEXT, COLOR_LINK)

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", self.CANCEL_CONFIRM):
            self._stop_event.set()
            self._pause_event.set()
            self.update_button_ui("stopping")
            self.set_status("■", self.STOPPING_STATUS_TEXT, COLOR_ERROR)

    # ---------- 树形表格通用操作 ----------
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
        if not selected_items:
            return
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
        self.set_status("✖", f"已剔除选中{self.ITEM_LABEL}。当前列表剩余 {len(self.tree.get_children())} {self.ITEM_UNIT}。",
                        COLOR_MUTED, COLOR_INFO)

    def on_tree_double_click(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        item_id = selected_items[0]
        out_path = self.item_outpath_map.get(item_id)
        if not out_path or not os.path.exists(out_path):
            out_path = self.item_filepath_map.get(item_id)
        if out_path and os.path.exists(out_path):
            try:
                os.startfile(out_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开位置:\n{str(e)}")

    def select_target(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path:
            self.target_var.set(path)

    def open_output_dir(self):
        tgt = self.target_var.get().strip()
        if os.path.exists(tgt):
            os.startfile(tgt)

    def update_tree_row(self, item_id, values, tag, auto_scroll=True):
        def update():
            if self.tree.exists(item_id):
                self.tree.item(item_id, values=values, tags=(tag,))
                if auto_scroll:
                    self.tree.see(item_id)

        self.parent.after(0, update)

    def clear_list(self):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", "任务正在运行，请先暂停！")
            return
        self.tree.delete(*self.tree.get_children())
        self.item_filepath_map.clear()
        self.item_outpath_map.clear()
        self.set_status("▤", "列表已清空。您可以点击按钮或直接拖拽文件进来。", COLOR_MUTED, COLOR_INFO)

    # ---------- 拖拽 / 添加文件（参数化文件过滤与列构建） ----------
    def is_valid_file(self, path):
        raise NotImplementedError

    def build_row_values(self, idx, path):
        raise NotImplementedError

    def on_drag_drop(self, event):
        if self.is_running and not self._pause_event.is_set():
            messagebox.showwarning("提示", self.DRAG_RUNNING_WARN)
            return
        dropped_paths = self.tree.tk.splitlist(event.data)
        found = []
        for path in dropped_paths:
            if os.path.isfile(path) and self.is_valid_file(path):
                found.append(path)
            elif os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for f in files:
                        full = os.path.join(root, f)
                        if self.is_valid_file(full):
                            found.append(full)
        if found:
            self._append_to_queue(found)
            if not self.target_var.get():
                self.target_var.set(os.path.dirname(found[0]))
        else:
            messagebox.showinfo("提示", self.DROP_EMPTY_MSG)

    def add_files(self):
        paths = filedialog.askopenfilenames(title=self.ADD_FILES_TITLE, filetypes=self.ADD_FILES_FILETYPES)
        if paths:
            valid = [p for p in paths if self.is_valid_file(p)]
            self._append_to_queue(valid)
            if valid and not self.target_var.get():
                self.target_var.set(os.path.dirname(valid[0]))

    def add_folder(self):
        path = filedialog.askdirectory(title=self.ADD_FOLDER_TITLE)
        if path:
            found = []
            for root, dirs, files in os.walk(path):
                for f in files:
                    full = os.path.join(root, f)
                    if self.is_valid_file(full):
                        found.append(full)
            if found:
                self._append_to_queue(found)
                if not self.target_var.get():
                    self.target_var.set(path)
            else:
                messagebox.showinfo("扫描结果", self.ADD_FOLDER_EMPTY_MSG)

    def _append_to_queue(self, paths):
        existing_paths = list(self.item_filepath_map.values())
        added_count = 0
        for path in paths:
            if path not in existing_paths:
                idx = len(self.tree.get_children()) + 1
                values = self.build_row_values(idx, path)
                item_id = self.tree.insert("", tk.END, values=values, tags=("pending",))
                self.item_filepath_map[item_id] = path
                added_count += 1
        self.set_status("✔", f"成功添加 {added_count} {self.ITEM_UNIT}{self.ITEM_LABEL}。"
                             f"队列中共 {len(self.tree.get_children())} {self.ITEM_UNIT}。", COLOR_SUCCESS)
