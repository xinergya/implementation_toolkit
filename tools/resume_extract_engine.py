import os
import io
import re
import time
import shutil
import threading
import datetime
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    import pandas as pd
except ImportError:
    pd = None

# 核心：操作 Windows 系统的底层库
import win32com.client
import pythoncom


class ResumeExtractUI:
    """基于 Excel 数据驱动的 Word 逆向精准提取引擎 (V2.0 工业级重构版：沙箱+物理映射+无损剪裁)"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.word_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.target_var = tk.StringVar()

        self.search_tpl_var = tk.StringVar(value="{人员级别}{序号}")
        self.name_tpl_var = tk.StringVar(value="{序号}-{工号}-{姓名}-{岗位}")

        self.excel_columns = []

        self.anim_running = False
        self.anim_frame = 0
        self.spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

        self.setup_ui()

    # ================= UI 布局 =================
    def setup_ui(self):
        if pd is None:
            tk.Label(self.parent, text="⚠️ 缺少核心运行库！请在命令行运行：pip install pandas openpyxl",
                     fg="red", font=("Microsoft YaHei UI", 12, "bold")).pack(pady=50)
            return

        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(7, weight=1)

        # 1. 数据源区
        tk.Label(frame, text="1. 源文件 (Word):", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.word_var).grid(row=0, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Button(frame, text="选择 Word...", command=self.select_word, width=12).grid(row=0, column=2, pady=5)

        tk.Label(frame, text="2. 人员清单表 (Excel):", font=("Microsoft YaHei UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        entry_excel = tk.Entry(frame, textvariable=self.excel_var)
        entry_excel.grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        # 回车触发后台异步解析，防 UI 假死
        entry_excel.bind("<Return>", lambda e: self._parse_excel_headers_bg(self.excel_var.get().strip().strip('"').strip("'")))
        tk.Button(frame, text="选择 Excel...", command=self.select_excel, width=12).grid(row=1, column=2, pady=5)

        self.lbl_columns = tk.Text(frame, font=("Microsoft YaHei UI", 10), fg="#198754", bg="#F0F2F5",
                                   height=3, wrap=tk.WORD, bd=0, highlightthickness=0)
        self.lbl_columns.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=10, pady=2)
        self.lbl_columns.insert(tk.END, "💡 导入或粘贴 Excel 路径并回车后，此处将显示【可直接框选复制】的 {变量名}。")
        self.lbl_columns.bind("<Key>", lambda e: "break" if not (e.state & 4 and e.keysym.lower() == 'c') else None)

        # 2. 规则定制区
        tk.Label(frame, text="3. 逆向检索公式:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=3, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.search_tpl_var, font=("Microsoft YaHei UI", 11)).grid(row=3, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Label(frame, text="* 引擎自动穿透【目录/表格】排雷", font=("Microsoft YaHei UI", 9), fg="#D35400").grid(row=3, column=2, sticky=tk.W)

        tk.Label(frame, text="4. 智能重命名公式:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.name_tpl_var, font=("Microsoft YaHei UI", 11)).grid(row=4, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Label(frame, text="* 提取出来后叫什么名字", font=("Microsoft YaHei UI", 9), fg="#6C757D").grid(row=4, column=2, sticky=tk.W)

        # 3. 输出目录区
        tk.Label(frame, text="5. 简历保存至:", font=("Microsoft YaHei UI", 10, "bold")).grid(row=5, column=0, sticky=tk.W, pady=5)
        tk.Entry(frame, textvariable=self.target_var).grid(row=5, column=1, sticky=tk.EW, padx=10, pady=5, ipady=3)
        tk.Button(frame, text="浏览目录...", command=self.select_target, width=12).grid(row=5, column=2, pady=5)

        # 4. 控制面板
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)
        self.btn_pause = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)
        self.btn_cancel = tk.Button(btn_frame, font=("Microsoft YaHei UI", 11, "bold"))
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 5. 日志区
        self.log_area = scrolledtext.ScrolledText(frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
                                                  relief=tk.FLAT, padx=12, pady=12, state=tk.DISABLED)
        self.log_area.grid(row=7, column=0, columnspan=3, sticky=tk.NSEW, pady=5)

        self.log_area.tag_config("INFO", foreground="#2C3E50")
        self.log_area.tag_config("SUCCESS", foreground="#198754", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("WARNING", foreground="#D35400", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("ERROR", foreground="#DC3545", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("MAGENTA", foreground="#8E44AD", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("LINK", foreground="#0078D7", underline=True)
        self.log_area.tag_bind("LINK", "<Enter>", lambda e: self.log_area.config(cursor="hand2"))
        self.log_area.tag_bind("LINK", "<Leave>", lambda e: self.log_area.config(cursor=""))

        self.log_to_ui("✨ 工业级逆向精准提取引擎 (无损切割版) 已就绪！", "MAGENTA")
        self.update_button_ui("idle")

    # ================= 状态机与动画 =================
    def _animate_btn(self):
        if not self.anim_running: return
        if self._pause_event.is_set():
            frame_char = self.spinner_frames[self.anim_frame % len(self.spinner_frames)]
            self.btn_start.config(text=f"{frame_char} 正在极速检索切割...")
            self.anim_frame += 1
        self.parent.after(100, self._animate_btn)

    def update_button_ui(self, state):
        color_start_active = "#8E44AD"
        color_pause_active = "#F39C12"
        color_stop_active = "#DC3545"
        color_disabled_bg = "#E9ECEF"
        color_disabled_fg = "#ADB5BD"
        color_enabled_gray_bg = "#E0E0E0"
        color_enabled_gray_fg = "#333333"

        if state == "idle":
            self.anim_running = False
            self.btn_start.config(bg=color_start_active, fg="white", text="🧲 开始智能提取", state=tk.NORMAL, command=self.start_extract_thread)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏸ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏹ 停止", state=tk.DISABLED)
        elif state == "running":
            self.btn_start.config(bg=color_start_active, fg="white", state=tk.NORMAL, command=lambda: None)
            self.btn_pause.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏸ 暂停", state=tk.NORMAL, command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏹ 停止", state=tk.NORMAL, command=self.cancel_processing)
            if not self.anim_running:
                self.anim_running = True
                self.anim_frame = 0
                self._animate_btn()
        elif state == "paused":
            self.btn_start.config(bg=color_enabled_gray_bg, fg="#888888", text="⚙️ 已挂起", state=tk.NORMAL, command=lambda: None)
            self.btn_pause.config(bg=color_pause_active, fg="white", text="▶ 继续提取", state=tk.NORMAL, command=self.toggle_pause)
            self.btn_cancel.config(bg=color_enabled_gray_bg, fg=color_enabled_gray_fg, text="⏹ 停止", state=tk.NORMAL, command=self.cancel_processing)
        elif state == "stopping":
            self.anim_running = False
            self.btn_start.config(bg=color_disabled_bg, fg=color_disabled_fg, text="🛑 正在中止...", state=tk.DISABLED)
            self.btn_pause.config(bg=color_disabled_bg, fg=color_disabled_fg, text="⏸ 暂停", state=tk.DISABLED)
            self.btn_cancel.config(bg=color_stop_active, fg="white", text="⏹ 正在清理...", state=tk.NORMAL, command=lambda: None)

    # ================= 交互响应与后台任务调度 =================
    def select_word(self):
        path = filedialog.askopenfilename(title="选择总标书 Word", filetypes=[("Word", "*.docx;*.doc")])
        if path:
            self.word_var.set(os.path.abspath(path))

            # 【优化】：拆除 if not 判断，强制每次都联动更新输出目录
            # 获取 Word 文件的纯名称（不带后缀）
            base_name = os.path.splitext(os.path.basename(path))[0]

            # 动态生成专属的输出文件夹，例如："提取出的简历_2024年度总标书"
            smart_target_dir = os.path.abspath(os.path.join(os.path.dirname(path), f"提取出的简历_{base_name}"))
            self.target_var.set(smart_target_dir)

    def select_excel(self):
        path = filedialog.askopenfilename(title="选择人员名单 Excel", filetypes=[("Excel", "*.xlsx;*.xls")])
        if path:
            self.excel_var.set(os.path.abspath(path))
            self._parse_excel_headers_bg(path)

    def _parse_excel_headers_bg(self, path):
        """架构优化：将可能阻塞 UI 的 Excel 解析丢入后台线程"""
        if not path or not os.path.exists(path): return
        self._update_column_hint("⏳ 正在后台解析 Excel 表头，请稍候...")
        threading.Thread(target=self._parse_excel_headers_thread, args=(path,), daemon=True).start()

    def _parse_excel_headers_thread(self, path):
        try:
            # 使用全内存加载，防止读取时被锁定报错
            with open(path, 'rb') as f:
                excel_bytes = f.read()
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))

            target_sheet = '实施人员清单'
            sheet_to_read = target_sheet if target_sheet in xls.sheet_names else 0

            raw_header_row = pd.read_excel(xls, sheet_name=sheet_to_read, header=None, nrows=1).values[0]
            raw_headers = [str(x).strip() for x in raw_header_row if pd.notna(x) and str(x).strip() != '']

            seen, duplicates = set(), set()
            for h in raw_headers:
                if h in seen: duplicates.add(h)
                seen.add(h)

            if duplicates:
                self.parent.after(0, lambda: self._update_column_hint(f"⚠️ 发现重复表头: {', '.join(duplicates)}。请在 Excel 中修改以防数据错位！"))
                self.excel_columns = []
            else:
                self.excel_columns = raw_headers
                hint = "💡 成功读取！请框选复制下方变量填入公式:\n   " + "   ".join([f"{{{col}}}" for col in self.excel_columns])
                self.parent.after(0, lambda: self._update_column_hint(hint))
        except Exception as e:
            self.parent.after(0, lambda: self._update_column_hint(f"⚠️ 解析失败: {str(e)}"))

    def _update_column_hint(self, msg):
        self.lbl_columns.config(state=tk.NORMAL)
        self.lbl_columns.delete(1.0, tk.END)
        self.lbl_columns.insert(tk.END, msg)
        self.lbl_columns.config(state=tk.DISABLED)

    def select_target(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path: self.target_var.set(os.path.abspath(path))

    def log_to_ui(self, message, level="INFO", link_path=None, auto_scroll=True):
        def update():
            self.log_area.config(state=tk.NORMAL)
            timestamp = datetime.datetime.now().strftime("[%H:%M:%S] ")
            self.log_area.insert(tk.END, f"[{timestamp}] {message}\n", level)
            if link_path:
                tag_name = f"link_{datetime.datetime.now().timestamp()}"
                self.log_area.tag_config(tag_name, foreground="#0078D7", underline=True, font=("Microsoft YaHei UI", 10, "bold"))
                self.log_area.tag_bind(tag_name, "<Enter>", lambda e: self.log_area.config(cursor="hand2"))
                self.log_area.tag_bind(tag_name, "<Leave>", lambda e: self.log_area.config(cursor=""))
                self.log_area.tag_bind(tag_name, "<Button-1>", lambda e, p=link_path: os.startfile(p))
                self.log_area.insert(tk.END, f" ➔ 点击打开查看结果\n", tag_name)
            if auto_scroll: self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED)

        self.parent.after(0, update)

    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.update_button_ui("paused")
            self.log_to_ui("⏸ 提取任务已挂起，等待恢复...", "WARNING")
        else:
            self._pause_event.set()
            self.update_button_ui("running")
            self.log_to_ui("▶ 提取任务已恢复。", "INFO")

    def cancel_processing(self):
        if messagebox.askyesno("确认停止", "确定要中止提取任务吗？\n已提取完成的简历将被保留。"):
            self._stop_event.set()
            self._pause_event.set()
            self.update_button_ui("stopping")
            self.log_to_ui("🛑 正在安全停止任务并清理临时缓存...", "WARNING", auto_scroll=False)

    def start_extract_thread(self):
        word_path = self.word_var.get().strip().strip('"').strip("'")
        excel_path = self.excel_var.get().strip().strip('"').strip("'")
        tgt_path = self.target_var.get().strip().strip('"').strip("'")
        search_tpl = self.search_tpl_var.get().strip()
        name_tpl = self.name_tpl_var.get().strip()

        if not all([word_path, excel_path, tgt_path, search_tpl, name_tpl]):
            messagebox.showwarning("提示", "请确保所有路径和公式已填写完整！")
            return
        if not self.excel_columns:
            messagebox.showwarning("提示", "请先正确读取包含有效表头的人员名单 Excel！")
            return

        self._stop_event.clear()
        self._pause_event.set()
        self.update_button_ui("running")

        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)

        threading.Thread(target=self.process_extraction, args=(word_path, excel_path, tgt_path, search_tpl, name_tpl), daemon=True).start()

    # ================= 辅助解析工具 =================
    def _sanitize_filename(self, filename):
        clean = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', filename)
        return re.sub(r'_+', '_', clean).strip()[:100]

    def _set_page_numbering(self, doc, start_page):
        try:
            is_first = True
            for sec in doc.Sections:
                for i in range(1, 4):
                    for obj in (sec.Headers(i), sec.Footers(i)):
                        try:
                            if is_first:
                                obj.PageNumbers.RestartNumberingAtSection = True
                                obj.PageNumbers.StartingNumber = int(start_page)
                            else:
                                obj.PageNumbers.RestartNumberingAtSection = False
                            obj.Range.Fields.Update()
                        except:
                            pass
                is_first = False
            doc.Repaginate()
        except:
            pass

    def _clear_header_line(self, doc):
        """彻底清除 Word 新文档自带的默认页眉横线 (底边框)"""
        try:
            for sec in doc.Sections:
                # 遍历三种页眉类型：1=主页眉, 2=首页页眉, 3=偶数页页眉
                for i in range(1, 4):
                    hdr = sec.Headers(i)
                    # -3 代表 wdBorderBottom (底边框常量)
                    # 0 代表 wdLineStyleNone (无边框常量)
                    try:
                        hdr.Range.ParagraphFormat.Borders(-3).LineStyle = 0
                    except:
                        pass
                    try:
                        hdr.Range.Borders(-3).LineStyle = 0
                    except:
                        pass
        except Exception:
            pass

    # ================= 核心：物理映射与无损提取架构 =================
    def process_extraction(self, word_path, excel_path, tgt_path, search_tpl, name_tpl):
        # 严格的 COM 线程初始化
        pythoncom.CoInitialize()
        word_app = None
        master_doc = None
        temp_dir = ""
        start_time = time.time()

        try:
            word_path = os.path.abspath(word_path)
            excel_path = os.path.abspath(excel_path)
            tgt_path = os.path.abspath(tgt_path)

            if not os.path.exists(tgt_path): os.makedirs(tgt_path)

            self.log_to_ui("📊 [阶段 1] 正在加载数据字典...", "MAGENTA")

            # 使用内存读取防止 Excel 被锁定
            with open(excel_path, 'rb') as f:
                excel_bytes = f.read()
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))

            target_sheet = '实施人员清单'
            sheet_to_read = target_sheet if target_sheet in xls.sheet_names else 0
            df = pd.read_excel(xls, sheet_name=sheet_to_read, dtype=str).fillna("")
            df.columns = [str(col).strip() for col in df.columns]

            tasks = []
            for index, row in df.iterrows():
                search_text = search_tpl
                name_text = name_tpl
                for col in df.columns:
                    val = str(row[col]).strip()
                    search_text = search_text.replace(f"{{{col}}}", val)
                    name_text = name_text.replace(f"{{{col}}}", val)
                tasks.append({'search_word': search_text, 'file_name': name_text})

            self.log_to_ui(f"  └─ 成功生成 {len(tasks)} 条提取任务。")
            self.log_to_ui(f"\n🗺️ [阶段 2] 启动空间雷达：构建物理映射地图...", "MAGENTA")

            # 【修复点 1】手动创建系统临时沙箱，脱离 with 语法的死板管控
            temp_dir = tempfile.mkdtemp()
            temp_master_path = os.path.join(temp_dir, "~master_copy.docx")
            shutil.copy2(word_path, temp_master_path)

            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            word_app.ScreenUpdating = False

            # 只读模式打开，防止误操作
            master_doc = word_app.Documents.Open(temp_master_path, ReadOnly=True, Visible=False)

            raw_found_tasks = []
            keyword_occurrences = {}

            for t in tasks:
                if self._stop_event.is_set(): break

                word = t['search_word']
                occ = keyword_occurrences.get(word, 0) + 1
                keyword_occurrences[word] = occ

                search_range = master_doc.Content
                search_range.Find.ClearFormatting()

                valid_count = 0
                found_start_pos = -1
                found_page = 1
                word_clean = "".join(word.split())

                while search_range.Find.Execute(FindText=word, Forward=True, Wrap=0):  # 0 = wdFindStop
                    is_valid = True

                    if search_range.Information(12):  # 在表格中
                        is_valid = False
                    else:
                        # 目录检查
                        for toc in master_doc.TablesOfContents:
                            if search_range.Start >= toc.Range.Start and search_range.End <= toc.Range.End:
                                is_valid = False
                                break

                        if is_valid:
                            para_text = search_range.Paragraphs(1).Range.Text
                            para_clean = "".join(para_text.split()).replace('\r', '').replace('\x07', '')
                            if para_clean != word_clean:  # 防止数字子串命中
                                is_valid = False

                    if is_valid:
                        valid_count += 1
                        if valid_count == occ:
                            found_start_pos = search_range.Start
                            found_page = search_range.Information(10)  # wdActiveEndPageNumber
                            break

                    # 安全向前步进，防止死循环
                    search_range.Collapse(Direction=0)
                    search_range.End = master_doc.Content.End

                if found_start_pos != -1:
                    t['start_pos'] = found_start_pos
                    t['page'] = found_page
                    raw_found_tasks.append(t)
                    self.log_to_ui(f"  └─ 锁定目标: [{word}] (原件页码:{found_page})")
                else:
                    self.log_to_ui(f"  └─ ⚠️ 遗失: 未能在正文中找到 [{word}]", "WARNING")

            if self._stop_event.is_set() or len(raw_found_tasks) == 0:
                self.log_to_ui("未找到可提取的数据，任务中止。", "WARNING", auto_scroll=False)
                return

            # 严格基于物理坐标重新排序
            raw_found_tasks.sort(key=lambda x: x['start_pos'])

            # 去重
            valid_tasks = []
            seen_pos = set()
            for t in raw_found_tasks:
                if t['start_pos'] not in seen_pos:
                    seen_pos.add(t['start_pos'])
                    valid_tasks.append(t)

            # 依据物理排序计算每个区块的 End 位置
            for i in range(len(valid_tasks)):
                if i < len(valid_tasks) - 1:
                    valid_tasks[i]['end_pos'] = valid_tasks[i + 1]['start_pos']
                else:
                    valid_tasks[i]['end_pos'] = master_doc.Content.End

            self.log_to_ui(f"\n✂️ [阶段 3] 坐标锁死，启动【无损镜像提取】...", "MAGENTA")

            success_count = 0
            ext = os.path.splitext(word_path)[1]

            for i, t in enumerate(valid_tasks):
                if self._stop_event.is_set():
                    self.log_to_ui(f"❌ 已取消剩余 {len(valid_tasks) - i} 份简历提取...", "ERROR", auto_scroll=False)
                    break
                self._pause_event.wait()

                safe_name = self._sanitize_filename(t['file_name'])
                new_filename = f"{safe_name}{ext}"
                new_filepath = os.path.abspath(os.path.join(tgt_path, new_filename))

                self.log_to_ui(f"[{i + 1}/{len(valid_tasks)}] 镜像复制: {new_filename} ...")

                new_doc = None
                try:
                    # 使用 Range Copy + PasteAndFormat
                    source_range = master_doc.Range(Start=t['start_pos'], End=t['end_pos'])
                    source_range.Copy()

                    # COM 操作时由于剪贴板占用，加一个极短休眠防止锁死
                    time.sleep(0.05)

                    new_doc = word_app.Documents.Add(Visible=False)
                    # 16 = wdFormatOriginalFormatting (保留原样格式)
                    new_doc.Range(0, 0).PasteAndFormat(16)

                    self._set_page_numbering(new_doc, t.get('page', 1))

                    # ⬇️ 【新增】：在这里呼叫“横线橡皮擦”
                    self._clear_header_line(new_doc)
                    # ⬆️ ==========================

                    if os.path.exists(new_filepath): os.remove(new_filepath)
                    new_doc.SaveAs(new_filepath)
                    success_count += 1

                except Exception as e:
                    self.log_to_ui(f"  └─ ❌ 提取失败: {str(e)}", "ERROR")
                finally:
                    if new_doc and hasattr(new_doc, 'Close'):
                        try:
                            new_doc.Close(SaveChanges=False)
                        except:
                            pass

            if not self._stop_event.is_set():
                elapsed = time.time() - start_time
                self.log_to_ui(f"\n🎉 逆向提取圆满结束！耗时: {elapsed:.2f} 秒。\n成功无损剥离 {success_count} 份简历。", "SUCCESS", link_path=tgt_path)

        except Exception as global_e:
            error_msg = str(global_e)
            if "-2146959355" in error_msg or "服务器运行失败" in error_msg or "80080005" in error_msg:
                self.log_to_ui("❌ Word底层卡死！请在任务管理器结束所有 WINWORD.EXE 进程。", "ERROR")
            elif "800706BA" in error_msg or "RPC" in error_msg:
                self.log_to_ui("❌ Word引擎崩溃！可能遇到极度庞大或损坏的表格。", "ERROR")
            elif "Permission denied" in error_msg:
                self.log_to_ui("❌ 权限拒绝！要保存或覆盖的文件正被另一个软件打开。", "ERROR")
            else:
                self.log_to_ui(f"⚠️ 发生严重错误: {error_msg}", "ERROR")

        finally:
            # 1. 率先切断 COM 连接，下发关闭指令
            if master_doc is not None:
                try:
                    if hasattr(master_doc, 'Close'):
                        master_doc.Close(SaveChanges=False)
                except Exception:
                    pass
                finally:
                    master_doc = None

            if word_app is not None:
                try:
                    if hasattr(word_app, 'ScreenUpdating'): word_app.ScreenUpdating = True
                    if hasattr(word_app, 'Quit'): word_app.Quit()
                except Exception:
                    pass
                finally:
                    word_app = None

            pythoncom.CoUninitialize()

            # 【修复点 2】：强制休眠 0.5 秒，给底层的 WINWORD.EXE 进程一点时间去释放文件锁
            time.sleep(0.5)

            # 【修复点 3】：使用静默抹除法，无视一切占用报错，还你一个清爽的完结日志
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

            def reset_ui():
                self.update_button_ui("idle")

            self.parent.after(0, reset_ui)