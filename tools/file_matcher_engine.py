import os
import re
import shutil
import threading
import datetime
import hashlib
import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from thefuzz import fuzz

# ================= 全局常量配置 =================
VALID_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png', '.gif', '.jfif', '.bmp', '.heic'}
DEFAULT_KEYWORDS = "证书, 毕业, 学位, 技能, 资质, 身份, 高级, 中级, 学信网, 学历"
CONFIG_FILE = "file_matcher_config.json"


class AutoFileProcessorUI:
    """文件智能匹配分发工具（V2.0 智能感知与双轨匹配重构版）"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        # 线程控制事件
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.link_counter = 0
        self.current_keywords = self.load_config()

        # UI 变量绑定
        self.excel_var = tk.StringVar(self.parent)
        self.source_var = tk.StringVar(self.parent)
        self.target_var = tk.StringVar(self.parent)
        self.keyword_var = tk.StringVar(self.parent, value=self.current_keywords)
        self.match_mode_var = tk.StringVar(self.parent, value="OR")  # 默认使用宽松模式

        self.setup_ui()

    # ================= 配置文件读写逻辑 =================
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get("keywords", DEFAULT_KEYWORDS)
            except Exception:
                pass
        return DEFAULT_KEYWORDS

    def save_config(self, keywords_str):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({"keywords": keywords_str}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log_to_ui(f"⚠️ 配置文件保存失败: {str(e)}", "WARNING")

    # ================= 适配工作台的 UI 布局 =================
    def setup_ui(self):
        frame = tk.Frame(self.parent, padx=20, pady=20, bg="#FFFFFF")
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(7, weight=1)  # 日志区域自适应拉伸向下移动一行

        # 1. Excel 数据源
        tk.Label(frame, text="1. Excel 数据源:", font=("Microsoft YaHei UI", 10), bg="#FFFFFF").grid(row=0, column=0,
                                                                                                     sticky=tk.W,
                                                                                                     pady=5)
        tk.Entry(frame, textvariable=self.excel_var, font=("Microsoft YaHei UI", 10)).grid(row=0, column=1,
                                                                                           sticky=tk.EW, padx=10,
                                                                                           pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_excel, width=10).grid(row=0, column=2, pady=5)

        # 2. 源文件根目录
        tk.Label(frame, text="2. 源文件根目录:", font=("Microsoft YaHei UI", 10), bg="#FFFFFF").grid(row=1, column=0,
                                                                                                     sticky=tk.W,
                                                                                                     pady=5)
        tk.Entry(frame, textvariable=self.source_var, font=("Microsoft YaHei UI", 10)).grid(row=1, column=1,
                                                                                            sticky=tk.EW, padx=10,
                                                                                            pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_source, width=10).grid(row=1, column=2, pady=5)

        # 3. 输出保存位置
        tk.Label(frame, text="3. 输出保存位置:", font=("Microsoft YaHei UI", 10), bg="#FFFFFF").grid(row=2, column=0,
                                                                                                     sticky=tk.W,
                                                                                                     pady=5)
        tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10)).grid(row=2, column=1,
                                                                                            sticky=tk.EW, padx=10,
                                                                                            pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_target, width=10).grid(row=2, column=2, pady=5)

        # 4. 检索关键字
        tk.Label(frame, text="4. 检索关键字:", font=("Microsoft YaHei UI", 10), bg="#FFFFFF").grid(row=3, column=0,
                                                                                                   sticky=tk.NW,
                                                                                                   pady=(12, 5))
        tk.Entry(frame, textvariable=self.keyword_var, font=("Microsoft YaHei UI", 10)).grid(row=3, column=1,
                                                                                             columnspan=2, sticky=tk.EW,
                                                                                             padx=10, pady=(12, 5),
                                                                                             ipady=3)
        tk.Label(frame, text="* 多个关键字请用逗号分隔，支持模糊匹配。", font=("Microsoft YaHei UI", 9), fg="#6C757D",
                 bg="#FFFFFF").grid(row=4, column=1, sticky=tk.W, padx=10)

        # 5. 【新增】匹配模式策略
        tk.Label(frame, text="5. 匹配策略:", font=("Microsoft YaHei UI", 10, "bold"), bg="#FFFFFF", fg="#0078D7").grid(
            row=5, column=0, sticky=tk.W, pady=10)
        mode_frame = tk.Frame(frame, bg="#FFFFFF")
        mode_frame.grid(row=5, column=1, columnspan=2, sticky=tk.W, padx=5, pady=10)

        tk.Radiobutton(mode_frame, text="宽松模式：(工号 或 姓名) + 关键字", variable=self.match_mode_var, value="OR",
                       font=("Microsoft YaHei UI", 10), bg="#FFFFFF", activebackground="#FFFFFF").pack(side=tk.LEFT,
                                                                                                       padx=5)
        tk.Radiobutton(mode_frame, text="严格模式：(工号 与 姓名) + 关键字", variable=self.match_mode_var, value="AND",
                       font=("Microsoft YaHei UI", 10), bg="#FFFFFF", activebackground="#FFFFFF").pack(side=tk.LEFT,
                                                                                                       padx=15)

        # 6. 控制按钮区
        btn_frame = tk.Frame(frame, bg="#FFFFFF")
        btn_frame.grid(row=6, column=0, columnspan=3, pady=15, sticky=tk.EW)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        self.btn_start = tk.Button(btn_frame, text="▶ 开始智能分发", font=("Microsoft YaHei UI", 11, "bold"),
                                   bg="#0078D7", fg="white", command=self.start_processing)
        self.btn_start.grid(row=0, column=0, padx=5, sticky=tk.EW, ipady=5)

        self.btn_pause = tk.Button(btn_frame, text="⏸ 暂停", font=("Microsoft YaHei UI", 11), command=self.toggle_pause,
                                   state=tk.DISABLED)
        self.btn_pause.grid(row=0, column=1, padx=5, sticky=tk.EW, ipady=5)

        self.btn_cancel = tk.Button(btn_frame, text="⏹ 取消", font=("Microsoft YaHei UI", 11),
                                    command=self.cancel_processing, state=tk.DISABLED)
        self.btn_cancel.grid(row=0, column=2, padx=5, sticky=tk.EW, ipady=5)

        # 7. 富文本日志区
        self.log_area = scrolledtext.ScrolledText(
            frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
            relief=tk.FLAT, padx=12, pady=12, state=tk.DISABLED
        )
        self.log_area.grid(row=7, column=0, columnspan=3, sticky=tk.NSEW, pady=5)

        # 统一的富文本标签配置
        self.log_area.tag_config("INFO", foreground="#2C3E50")
        self.log_area.tag_config("SUCCESS", foreground="#198754", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("WARNING", foreground="#D35400", font=("Microsoft YaHei UI", 10, "bold"))
        self.log_area.tag_config("ERROR", foreground="#DC3545", font=("Microsoft YaHei UI", 10, "bold"))

    # ================= UI 交互与线程控制 =================
    def select_excel(self):
        filepath = filedialog.askopenfilename(title="选择包含员工信息的 Excel",
                                              filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filepath:
            self.excel_var.set(filepath)
            self.target_var.set(os.path.join(os.path.dirname(filepath), "归档输出目录"))

    def select_source(self):
        dirpath = filedialog.askdirectory(title="选择源文件根目录")
        if dirpath: self.source_var.set(dirpath)

    def select_target(self):
        dirpath = filedialog.askdirectory(title="选择输出保存位置")
        if dirpath: self.target_var.set(dirpath)

    def log_to_ui(self, message, level="INFO", hyperlink_path=None):
        def update_ui():
            self.log_area.config(state=tk.NORMAL)
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_area.insert(tk.END, f"[{timestamp}] {message}\n", level)

            if hyperlink_path:
                self.link_counter += 1
                link_tag = f"link_{self.link_counter}"
                self.log_area.tag_config(link_tag, foreground="#0078D7", underline=True,
                                         font=("Microsoft YaHei UI", 10, "bold"))
                self.log_area.tag_bind(link_tag, "<Enter>", lambda e: self.log_area.config(cursor="hand2"))
                self.log_area.tag_bind(link_tag, "<Leave>", lambda e: self.log_area.config(cursor=""))
                self.log_area.tag_bind(link_tag, "<Button-1>", lambda e, p=hyperlink_path: os.startfile(p))
                self.log_area.insert(tk.END, f" ➔ 点击立即打开: {hyperlink_path}\n", link_tag)

            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED)

        self.parent.after(0, update_ui)

    def toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self.btn_pause.config(text="▶ 继续")
            self.log_to_ui("⏸ 处理已挂起，等待恢复...", "WARNING")
        else:
            self._pause_event.set()
            self.btn_pause.config(text="⏸ 暂停")
            self.log_to_ui("▶ 处理已恢复。", "INFO")

    def cancel_processing(self):
        if messagebox.askyesno("确认取消", "确定要中止当前的文件处理任务吗？\n已处理的文件将保留。"):
            self.log_to_ui("⚠️ 正在安全停止后台任务，请稍候...", "WARNING")
            self._stop_event.set()
            self._pause_event.set()
            self.btn_pause.config(state=tk.DISABLED)
            self.btn_cancel.config(state=tk.DISABLED)

    def start_processing(self):
        excel_path = self.excel_var.get().strip().strip('"').strip("'")
        source_dir = self.source_var.get().strip().strip('"').strip("'")
        target_dir = self.target_var.get().strip().strip('"').strip("'")
        raw_keywords = self.keyword_var.get().strip()
        match_mode = self.match_mode_var.get()

        if not raw_keywords:
            messagebox.showwarning("提示", "检索关键字不能为空！")
            return

        keywords_list = [k.strip() for k in re.split(r'[,，、]', raw_keywords) if k.strip()]

        if not all([excel_path, source_dir, target_dir]):
            messagebox.showwarning("提示", "请完整填写 Excel路径、源目录和输出目录！")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("错误", "Excel 数据源文件不存在！")
            return

        self.save_config(raw_keywords)
        self._stop_event.clear()
        self._pause_event.set()

        self.btn_start.config(state=tk.DISABLED, text="⚙ 处理中...")
        self.btn_pause.config(state=tk.NORMAL, text="⏸ 暂停")
        self.btn_cancel.config(state=tk.NORMAL)

        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)

        threading.Thread(target=self.process_data, args=(excel_path, source_dir, target_dir, keywords_list, match_mode),
                         daemon=True).start()

    # ================= 核心算法 =================
    def _clean_filename(self, text):
        return re.sub(r'[\\/:*?"<>|\s]', '_', str(text).strip())

    def _clean_feature_string(self, text):
        if not text: return ""
        return re.sub(r'[^\w\u4e00-\u9fa5]', '', str(text).lower())

    def _find_target_columns(self, df):
        """【新增】智能嗅探 Excel 表头位置"""
        id_col, name_col = None, None

        # 定义常见表头变体
        id_keywords = ['工号', '编号', '员工编码', '人员编号', 'id']
        name_keywords = ['姓名', '名字', '员工名称']

        for col in df.columns:
            col_str = str(col).lower()
            if not id_col and any(k in col_str for k in id_keywords):
                id_col = col
            if not name_col and any(k in col_str for k in name_keywords):
                name_col = col

        return id_col, name_col

    def _is_file_matched(self, original_filename, emp_name, emp_id, keywords_list, match_mode):
        """【修复】基于原始文件名的独立鉴权引擎"""
        # 仅针对文件名本身进行清理，杜绝父级目录特征污染
        clean_filename = self._clean_feature_string(original_filename)
        clean_name = self._clean_feature_string(emp_name)
        clean_id = self._clean_feature_string(emp_id)

        # 身份校验：严格在文件名范围内查找
        match_id = bool(clean_id and clean_id in clean_filename)
        match_name = bool(clean_name and clean_name in clean_filename)

        # 身份鉴权逻辑
        if match_mode == "AND":
            has_identity = match_id and match_name
        else:  # "OR" 宽松模式
            has_identity = match_id or match_name

        if not has_identity:
            return False

        # 关键字鉴权逻辑：优先进行 O(1) 的精准子串匹配，再使用 Fuzz 兜底
        for kw in keywords_list:
            clean_kw = self._clean_feature_string(kw)
            if not clean_kw:
                continue

            # 优化：精准包含直接返回 True，否则计算相似度
            if clean_kw in clean_filename or fuzz.partial_ratio(clean_kw, clean_filename) >= 80:
                return True

        return False

    def _get_file_md5(self, filepath):
        hasher = hashlib.md5()
        try:
            with open(filepath, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception:
            return None

    def _check_keywords(self, clean_filename, keywords_list):
        """高速关键字拦截引擎"""
        for kw in keywords_list:
            clean_kw = self._clean_feature_string(kw)
            if not clean_kw:
                continue
            # 优先使用底层 C 实现的 in 操作符，失败后再进行模糊计算兜底
            if clean_kw in clean_filename or fuzz.partial_ratio(clean_kw, clean_filename) >= 80:
                return True
        return False

    # ================= 主工作流 =================
    def process_data(self, excel_path, source_dir, target_dir, keywords_list, match_mode):
        mode_text = "严格模式 (工号+姓名)" if match_mode == "AND" else "宽松模式 (工号或姓名)"
        self.log_to_ui(f"🚀 启动极速处理引擎... | 策略: {mode_text} | 架构: 双轨 Hash 映射", "INFO")

        try:
            if not os.path.exists(target_dir): os.makedirs(target_dir)

            # ================= 阶段 1：解析 Excel，构建倒排索引 (O(N)) =================
            self.log_to_ui("📊 正在构建员工实体 Hash 索引库...", "INFO")
            df = pd.read_excel(excel_path, dtype=str)
            id_col, name_col = self._find_target_columns(df)

            if not id_col or not name_col:
                self.log_to_ui(f"❌ 无法在 Excel 中识别出【工号】或【姓名】列，请检查表头", "ERROR")
                return

            if "处理状态" not in df.columns: df["处理状态"] = ""
            if "落盘路径" not in df.columns: df["落盘路径"] = ""

            employees = {}  # 索引主存： row_index -> 员工档案字典
            id_map = {}  # 倒排索引： clean_id -> set(row_index)
            name_map = {}  # 倒排索引： clean_name -> set(row_index)

            for index, row in df.iterrows():
                # 【新增防御】处理 Pandas 读取 Excel 纯数字工号时自动补 ".0" 的幽灵 Bug
                emp_id = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
                if emp_id.endswith(".0"):
                    emp_id = emp_id[:-2]  # 将 "110385.0" 还原为 "110385"

                # 【防脏数据防御 2】：强力剔除 Excel 姓名中可能混入的空格或不可见字符
                emp_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                emp_name = re.sub(r'\s+', '', emp_name)

                if not emp_id and not emp_name: continue

                clean_id = self._clean_feature_string(emp_id)
                clean_name = self._clean_feature_string(emp_name)

                # 初始化员工对象
                employees[index] = {
                    "emp_id": emp_id,
                    "emp_name": emp_name,
                    "matched_files": [],  # 预留匹配队列
                    "folder_name": f"{index + 1}_{self._clean_filename(emp_id)}_{self._clean_filename(emp_name)}"
                }

                # 挂载 Hash 节点
                if clean_id: id_map.setdefault(clean_id, set()).add(index)
                if clean_name: name_map.setdefault(clean_name, set()).add(index)

            self.log_to_ui(f"✅ 索引构建完毕，共装载 {len(employees)} 条员工主数据。", "SUCCESS")

            # ================= 阶段 2：单次遍历磁盘，极速反向寻址 (O(M)) =================
            self.log_to_ui("📂 正在进行全局高速扫盘与特征解析...", "INFO")

            matched_relationship_count = 0
            for root_dir, _, files in os.walk(source_dir):
                if self._stop_event.is_set(): break
                self._pause_event.wait()

                for file_name in files:
                    ext = os.path.splitext(file_name)[1].lower()
                    if ext not in VALID_EXTENSIONS: continue

                    fpath = os.path.join(root_dir, file_name)
                    clean_fname = self._clean_feature_string(file_name)

                    # 【性能爆点 1】前置校验：如果文件不是目标资产（证书/资质等），直接抛弃
                    if not self._check_keywords(clean_fname, keywords_list):
                        continue

                    # 【性能爆点 2】提取潜在特征并触发 Hash 寻址
                    # 工号检索：提取文件名中的连续数字/字母串，去 id_map 中 O(1) 瞬间定位
                    # 维度 A：直接从原始文件名提取，保留天然边界（解决 98-110011 连字符合并问题）
                    # 【核心修复】：直接从原始 file_name 提取特征，保留连字符和空格的阻断作用！
                    # 这样 98-110011 就能被完美拆分为 '98' 和 '110011'
                    potential_ids = set(re.findall(r'[a-zA-Z0-9]+', file_name.lower()))
                    # 维度 B：作为兜底，也从无符号的干净文件名中提取（解决工号本身带特殊字符被拆断的问题）
                    potential_ids.update(re.findall(r'[a-zA-Z0-9]+', clean_fname))

                    matched_by_id = set()
                    for pid in potential_ids:
                        if pid in id_map:
                            matched_by_id.update(id_map[pid])

                    # 姓名检索：调用底层 C 实现的字符串包含检测
                    matched_by_name = set()
                    for cname, indices in name_map.items():
                        if cname in clean_fname:
                            matched_by_name.update(indices)

                    # 执行鉴权策略：将文件分发给对应的候选员工
                    if match_mode == "AND":
                        target_indices = matched_by_id & matched_by_name
                    else:  # OR 模式
                        target_indices = matched_by_id | matched_by_name

                    # 将文件压入对应员工的待处理队列
                    for idx in target_indices:
                        employees[idx]["matched_files"].append((file_name, fpath))
                        matched_relationship_count += 1

            if self._stop_event.is_set(): return
            self.log_to_ui(f"🎯 特征寻址完成！建立命中映射: {matched_relationship_count} 组。开始物理归档...", "INFO")

            # ================= 阶段 3：执行物理文件落盘 (I/O) =================
            success_count = 0
            for index, emp in employees.items():
                if self._stop_event.is_set(): break
                self._pause_event.wait()

                matched_files = emp["matched_files"]
                if not matched_files:
                    df.at[index, "处理状态"] = "⚠️ 未匹配到资质材料"
                    continue

                emp_target_dir = os.path.join(target_dir, emp["folder_name"])
                if not os.path.exists(emp_target_dir): os.makedirs(emp_target_dir)

                copied = 0
                seen_hashes = set()

                for fname, fpath in matched_files:
                    # 文件排重机制
                    file_md5 = self._get_file_md5(fpath)
                    if file_md5:
                        if file_md5 in seen_hashes: continue
                        seen_hashes.add(file_md5)

                    dst_path = os.path.join(emp_target_dir, fname)
                    base, target_ext = os.path.splitext(fname)
                    counter = 1
                    # 命名冲突处理
                    while os.path.exists(dst_path):
                        dst_path = os.path.join(emp_target_dir, f"{base}({counter}){target_ext}")
                        counter += 1

                    shutil.copy2(fpath, dst_path)
                    copied += 1

                df.at[index, "处理状态"] = f"✅ 已成功归档 {copied} 个文件"
                df.at[index, "落盘路径"] = emp_target_dir
                success_count += 1
                self.log_to_ui(f"✅ [{emp['emp_name']}] 归档 {copied} 个文件", "SUCCESS", hyperlink_path=emp_target_dir)

            # 输出结果报表
            result_excel_name = f"分发结果报表_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            result_excel_path = os.path.join(os.path.dirname(excel_path), result_excel_name)
            df.to_excel(result_excel_path, index=False)

            if not self._stop_event.is_set():
                self.log_to_ui(f"🎉 全部处理完毕！成功处理 {success_count} 名员工档案。", "SUCCESS")
            self.log_to_ui("📊 结果状态报表已生成", "INFO", hyperlink_path=result_excel_path)

        except Exception as e:
            self.log_to_ui(f"❌ 发生全局异常: {str(e)}", "ERROR")
        finally:
            def reset_ui():
                self.btn_start.config(state=tk.NORMAL, text="▶ 开始智能分发")
                self.btn_pause.config(state=tk.DISABLED, text="⏸ 暂停")
                self.btn_cancel.config(state=tk.DISABLED)

            self.parent.after(0, reset_ui)