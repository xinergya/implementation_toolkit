import os
import re
import time
import shutil
import threading
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# 核心：操作 Windows 系统的底层库
import win32com.client
import pythoncom


class WordSplitUI:
    """Word 超大文档极速拆分引擎 (单文件精细流 + 状态机 + 防滚日志)"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.level_var = tk.StringVar(value="1级标题")

        self.is_running = False

        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)  # 完美拉伸第5行的日志框

        # 1. 数据源选择 (支持直接粘贴并回车)
        tk.Label(frame, text="1. 待拆分的Word文档:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0,
                                                                                                   sticky=tk.W, pady=5)
        # 【体验优化】：去掉 readonly，允许用户直接 Ctrl+V 粘贴路径
        self.entry_source = tk.Entry(frame, textvariable=self.source_var, font=("Microsoft YaHei UI", 10, "bold"))
        self.entry_source.grid(row=0, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        self.btn_source = tk.Button(frame, text="选择文件...", command=self.select_source, width=12)
        self.btn_source.grid(row=0, column=2, pady=5)

        # 2. 输出目录选择
        tk.Label(frame, text="2. 子文档保存至:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0,
                                                                                               sticky=tk.W, pady=5)
        self.entry_target = tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10, "bold"))
        self.entry_target.grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.select_target, width=12).grid(row=1, column=2, pady=5)

        # 3. 拆分配置区
        tk.Label(frame, text="3. 拆分依据(标题级别):", font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0,
                                                                                                     sticky=tk.W,
                                                                                                     pady=5)
        level_frame = tk.Frame(frame)
        level_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=10, pady=5)

        # ====== 【新增：重写下拉框只读状态下的底色】 ======
        style = ttk.Style()
        style.map('TCombobox', fieldbackground=[('readonly', '#FFFFFF')])
        # =================================================

        self.level_combo = ttk.Combobox(level_frame, textvariable=self.level_var, font=("Microsoft YaHei UI", 10),
                                        width=15, state="readonly")
        self.level_combo['values'] = tuple(f"{i}级标题" for i in range(1, 10))
        self.level_combo.pack(side=tk.LEFT, ipady=3)
        tk.Label(level_frame, text="* 引擎将绝对锁死页码，确保子文档页码与原件完全一致", font=("Microsoft YaHei UI", 10),
                 fg="#6C757D").pack(side=tk.LEFT, padx=(10, 0))

        # 4. 控制按钮区 (引入顶级状态机架构)
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)

        self.btn_pause = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 5. 实时日志区
        tk.Label(frame, text="处理日志:", font=("Microsoft YaHei UI", 10, "bold"), fg="#333333").grid(row=4, column=0,
                                                                                                      sticky=tk.W,
                                                                                                      pady=(10, 5))

        self.log_area = scrolledtext.ScrolledText(
            frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
            relief=tk.FLAT, padx=12, pady=12, state=tk.DISABLED
        )
        self.log_area.grid(row=5, column=0, columnspan=3, sticky=tk.NSEW, pady=5)

        # 配置富文本颜色标签
        self.log_area.tag_config("INFO", foreground="#2C3E50")
        self.log_area.tag_config("SUCCESS", foreground="#198754", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("WARNING", foreground="#D35400", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("ERROR", foreground="#DC3545", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("MAGENTA", foreground="#8E44AD", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("LINK", foreground="#0078D7", underline=True, font=("Microsoft YaHei UI", 10, "bold"))

        self.log_area.tag_bind("LINK", "<Enter>", lambda e: self.log_area.config(cursor="hand2"))
        self.log_area.tag_bind("LINK", "<Leave>", lambda e: self.log_area.config(cursor=""))

        self.log_to_ui("✨ Word 单文件精细拆分引擎已就绪！")
        self.update_button_ui("idle")

    # ================= 按钮状态机管理器 =================
    def update_button_ui(self, state):
        color_start_active = "#2980B9"  # 经典蓝
        color_pause_active = "#F39C12"
        color_stop_active = "#DC3545"

        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray_bg = "#E0E0E0"
        color_enabled_gray_fg = "#333333"

        if state == "idle":
            self.btn_start.config(bg=color_start_active, fg="white", text="✂️ 开始智能拆分", state=tk.NORMAL,
                                  command=self.start_split_thread)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏸ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏹ 停止", state=tk.DISABLED)

        elif state == "running":
            self.btn_start.config(bg=color_start_active, fg="white", text="⚙️ 极速拆分中...", state=tk.NORMAL,
                                  command=lambda: None)
            self.btn_pause.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏸ 暂停", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏹ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)

        elif state == "paused":
            self.btn_start.config(bg=color_enabled_gray_bg, fg="#888888", text="⚙️ 已挂起", state=tk.NORMAL,
                                  command=lambda: None)
            self.btn_pause.config(bg=color_pause_active, fg="white", text="▶ 继续拆分", state=tk.NORMAL,
                                  command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏹ 停止", state=tk.NORMAL,
                                   command=self.cancel_processing)

        elif state == "stopping":
            self.btn_start.config(bg=color_disabled_bg, fg=color_disabled_fg, text="🛑 正在中止...", state=tk.DISABLED)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏸ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="⏹ 正在停止...", state=tk.NORMAL,
                                   command=lambda: None)

    # ================= UI 交互流 =================
    def select_source(self):
        path = filedialog.askopenfilename(title="请选择一个需要拆分的 Word 文档",
                                          filetypes=[("Word 文档", "*.docx;*.doc")])
        if path:
            self.source_var.set(path)
            base_name = os.path.splitext(os.path.basename(path))[0]
            self.target_var.set(os.path.join(os.path.dirname(path), f"{base_name}_拆分结果"))

    def select_target(self):
        path = filedialog.askdirectory(title="选择子文档保存目录")
        if path: self.target_var.set(path)

    # 【核心体验优化】：加入 auto_scroll 参数，防止点击停止后视线被强行拽到最下面
    def log_to_ui(self, message, level="INFO", link_path=None, auto_scroll=True):
        def update():
            self.log_area.config(state=tk.NORMAL)
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_area.insert(tk.END, f"[{timestamp}] {message}\n", level)
            if link_path:
                tag_name = f"link_{datetime.datetime.now().timestamp()}"
                self.log_area.tag_config(tag_name, foreground="#0078D7", underline=True,
                                         font=("Microsoft YaHei UI", 10, "bold"))
                self.log_area.tag_bind(tag_name, "<Enter>", lambda e: self.log_area.config(cursor="hand2"))
                self.log_area.tag_bind(tag_name, "<Leave>", lambda e: self.log_area.config(cursor=""))
                self.log_area.tag_bind(tag_name, "<Button-1>", lambda e, p=link_path: os.startfile(p))
                self.log_area.insert(tk.END, f" ➔ 点击打开存放目录\n", tag_name)

            if auto_scroll:
                self.log_area.see(tk.END)

            self.log_area.config(state=tk.DISABLED)

        self.parent.after(0, update)

    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.update_button_ui("paused")
            self.log_to_ui("⏸ 拆分任务已挂起，等待恢复...", "WARNING")
        else:
            self._pause_event.set()
            self.update_button_ui("running")
            self.log_to_ui("▶ 拆分任务已恢复。", "INFO")

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", "确定要中止当前的拆分任务吗？\n已经生成的子文档将被保留。"):
            self._stop_event.set()
            self._pause_event.set()
            self.update_button_ui("stopping")
            # 点击停止时，静默输入提示，不再强行滚动
            self.log_to_ui("🛑 正在安全停止任务，等待底层 Word 引擎释放资源...", "WARNING", auto_scroll=False)

    def start_split_thread(self):
        src = self.source_var.get().strip().strip('"').strip("'")
        tgt = self.target_var.get().strip().strip('"').strip("'")

        level_str = self.level_var.get()
        target_level = int(re.search(r'\d+', level_str).group())

        if not src or not tgt:
            messagebox.showwarning("提示", "请先输入/选择待拆分的 Word 文档以及输出目录！")
            return

        if not os.path.exists(src):
            messagebox.showerror("错误", "源文件不存在，请检查路径！")
            return

        # 【核心体验优化：目录防覆盖检查】
        if os.path.exists(tgt) and os.path.isdir(tgt) and os.listdir(tgt):
            if not messagebox.askyesno("防覆盖提示",
                                       "⚠️ 输出目录非空！\n\n重新开始可能会直接覆盖里面已存在的同名文档。\n是否确认覆盖并继续执行？"):
                return

        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        self.update_button_ui("running")

        # 每次重新开始时清空日志
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)

        threading.Thread(target=self.process_split, args=(src, tgt, target_level), daemon=True).start()

    # ================= 辅助方法 =================
    def _sanitize_filename(self, filename):
        clean = re.sub(r'[\\/:*?"<>|]', '_', filename)
        return clean[:50]

    def _safe_delete(self, doc, start_pos, end_pos):
        try:
            doc.Range(Start=start_pos, End=end_pos).Delete()
        except Exception:
            try:
                doc.Range(Start=start_pos, End=end_pos - 1).Delete()
            except Exception:
                doc.Range(Start=start_pos, End=end_pos).Text = ""

    def _set_page_numbering(self, doc, start_page):
        """强制突破 Word 排版重置的封锁，死锁真实页码并强制刷新所有域"""
        try:
            is_first = True
            for sec in doc.Sections:
                for i in range(1, 4):
                    try:
                        hdr = sec.Headers(i)
                        if is_first:
                            hdr.PageNumbers.RestartNumberingAtSection = True
                            hdr.PageNumbers.StartingNumber = int(start_page)
                        else:
                            hdr.PageNumbers.RestartNumberingAtSection = False
                        hdr.Range.Fields.Update()
                    except Exception:
                        pass
                    try:
                        ftr = sec.Footers(i)
                        if is_first:
                            ftr.PageNumbers.RestartNumberingAtSection = True
                            ftr.PageNumbers.StartingNumber = int(start_page)
                        else:
                            ftr.PageNumbers.RestartNumberingAtSection = False
                        ftr.Range.Fields.Update()
                    except Exception:
                        pass
                is_first = False
            doc.Repaginate()
        except Exception:
            pass

    # ================= 核心自动化拆分逻辑 =================
    def process_split(self, src, tgt, target_level):
        pythoncom.CoInitialize()
        word_app = None
        master_doc = None
        start_time = time.time()
        chapters = []
        temp_master_path = ""

        try:
            if not os.path.exists(tgt):
                os.makedirs(tgt)

            self.log_to_ui(f"🧹 [阶段 1] 正在创建并清洗母版文档 (解除域链接、接受修订)...", "MAGENTA")
            temp_master_path = os.path.join(tgt, "~temp_master_clean.docx")

            # 【修复覆盖被占用】：清道夫逻辑
            if os.path.exists(temp_master_path):
                try:
                    os.remove(temp_master_path)
                except PermissionError:
                    raise PermissionError("母版缓存文件正被其他程序打开")

            shutil.copy2(src, temp_master_path)

            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            word_app.ScreenUpdating = False

            master_doc = word_app.Documents.Open(temp_master_path, Visible=False)
            master_doc.Fields.Unlink()
            if master_doc.Revisions.Count > 0:
                master_doc.AcceptAllRevisions()
            master_doc.TrackRevisions = False
            master_doc.Save()

            self.log_to_ui(f"🔍 [阶段 2] 启动底层检索，收集【{target_level}级标题】坐标及物理页码...", "MAGENTA")
            master_doc.ActiveWindow.View.Type = 1

            find_obj = master_doc.Content.Find
            find_obj.ClearFormatting()
            find_obj.ParagraphFormat.OutlineLevel = target_level
            find_obj.Text = ""
            find_obj.Forward = True
            find_obj.Wrap = 0

            while find_obj.Execute():
                if self._stop_event.is_set():
                    break
                found_range = find_obj.Parent
                text = found_range.Text.replace('\r', '').replace('\x07', '').strip()

                if text:
                    page_num = found_range.Information(10)
                    chapters.append({
                        'text': text,
                        'page': page_num,
                        'start_pos': found_range.Start
                    })
                    self.log_to_ui(f"  └─ 锁定: [{text}] (原件第 {page_num} 页)")

                found_range.Collapse(Direction=0)

            master_doc.Close(False)
            master_doc = None

            if self._stop_event.is_set():
                self.log_to_ui("🛑 任务在扫描阶段被用户中止！", "WARNING", auto_scroll=False)
                return

            total_chapters = len(chapters)
            if total_chapters == 0:
                self.log_to_ui(f"❌ 失败: 文档中未发现该级别的标题", "ERROR")
                return

            self.log_to_ui(f"\n🚀 [阶段 3] 提取到 {total_chapters} 个稳定坐标，启动 [盲切与页码锁死] 引擎...", "MAGENTA")

            base_filename = os.path.splitext(os.path.basename(src))[0]
            ext = os.path.splitext(src)[1]
            success_count = 0

            for i, chapter in enumerate(chapters):
                # 中断时，把剩下还没切的记录标红为“取消”并防止屏幕疯狂滚动
                if self._stop_event.is_set():
                    remaining = total_chapters - i
                    self.log_to_ui(f"❌ 已取消剩余 {remaining} 个章节的拆分...", "ERROR", auto_scroll=False)
                    break
                self._pause_event.wait()

                self.log_to_ui(
                    f"[{i + 1}/{total_chapters}] 切割并锁定起始页为 {chapter['page']}: {chapter['text']} ...")

                work_doc = None
                try:
                    curr_start = chapter['start_pos']
                    next_start = chapters[i + 1]['start_pos'] if i + 1 < total_chapters else None

                    safe_title = self._sanitize_filename(chapter['text'])
                    new_filename = f"{i + 1:02d}_{safe_title}{ext}"
                    new_filepath = os.path.join(tgt, new_filename)

                    # 【优雅覆盖逻辑】：正式写入前先由 Python 扫清障碍
                    if os.path.exists(new_filepath):
                        os.remove(new_filepath)

                    shutil.copy2(temp_master_path, new_filepath)
                    work_doc = word_app.Documents.Open(new_filepath, Visible=False)

                    if next_start is not None:
                        self._safe_delete(work_doc, next_start, work_doc.Content.End)
                    if curr_start > 0:
                        self._safe_delete(work_doc, 0, curr_start)

                    # 核心防重排系统
                    self._set_page_numbering(work_doc, chapter['page'])

                    work_doc.Save()
                    success_count += 1
                    self.log_to_ui(f"  └─ ✅ 成功生成", "SUCCESS")

                # 精准友好的异常捕获
                except PermissionError:
                    self.log_to_ui(f"  └─ ❌ 覆盖失败: 文档被占用或无权限", "ERROR")
                except Exception as e:
                    error_msg = str(e)
                    if "80080005" in error_msg or "服务器运行失败" in error_msg:
                        self.log_to_ui(f"  └─ ❌ Word引擎卡死", "ERROR")
                    elif "800706BA" in error_msg or "RPC" in error_msg:
                        self.log_to_ui(f"  └─ ❌ Word底层崩溃闪退", "ERROR")
                    else:
                        self.log_to_ui(f"  └─ ❌ 生成失败: {error_msg}", "ERROR")
                finally:
                    if work_doc:
                        try:
                            work_doc.Close(False)
                        except:
                            pass

            if not self._stop_event.is_set():
                elapsed = time.time() - start_time
                self.log_to_ui(
                    f"\n🎉 极速拆分任务圆满结束！耗时: {elapsed:.2f} 秒。\n共成功剥离 {success_count} 个子文档。",
                    "SUCCESS", link_path=tgt)

        except Exception as global_e:
            error_msg = str(global_e)
            if "80080005" in error_msg or "服务器运行失败" in error_msg:
                self.log_to_ui("❌ 底层 Word 引擎启动失败！请按下 Ctrl+Shift+Esc 结束所有 WINWORD.EXE 进程。", "ERROR")
            else:
                self.log_to_ui(f"⚠️ 发生全局严重错误: {error_msg}", "ERROR")

        finally:
            # 安全清场
            if os.path.exists(temp_master_path):
                try:
                    os.remove(temp_master_path)
                except:
                    pass
            if master_doc:
                try:
                    master_doc.Close(False)
                except:
                    pass
            if word_app:
                try:
                    word_app.ScreenUpdating = True
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

            def reset_ui():
                self.is_running = False
                self.update_button_ui("idle")

            self.parent.after(0, reset_ui)