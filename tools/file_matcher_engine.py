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
    """文件智能匹配分发工具（已适配多标签页架构）"""

    def __init__(self, parent_frame):
        self.parent = parent_frame

        # 线程控制事件
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

        self.link_counter = 0
        self.current_keywords = self.load_config()

        # 变量绑定
        self.excel_var = tk.StringVar(self.parent)
        self.source_var = tk.StringVar(self.parent)
        self.target_var = tk.StringVar(self.parent)
        self.keyword_var = tk.StringVar(self.parent, value=self.current_keywords)

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
        # 抛弃原有的 main_frame，直接以传入的 parent 作为底层画布
        frame = tk.Frame(self.parent, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)  # 日志区域自适应拉伸

        # 1. Excel 数据源
        tk.Label(frame, text="1. Excel 数据源:", font=("Microsoft YaHei UI", 10)).grid(row=0, column=0, sticky=tk.W,
                                                                                       pady=5)
        tk.Entry(frame, textvariable=self.excel_var, font=("Microsoft YaHei UI", 10)).grid(row=0, column=1, sticky=tk.EW, padx=10,
                                                                            pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_excel, width=10).grid(row=0, column=2, pady=5)

        # 2. 源文件根目录
        tk.Label(frame, text="2. 源文件根目录:", font=("Microsoft YaHei UI", 10)).grid(row=1, column=0, sticky=tk.W,
                                                                                       pady=5)
        tk.Entry(frame, textvariable=self.source_var, font=("Microsoft YaHei UI", 10)).grid(row=1, column=1, sticky=tk.EW, padx=10,
                                                                             pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_source, width=10).grid(row=1, column=2, pady=5)

        # 3. 输出保存位置
        tk.Label(frame, text="3. 输出保存位置:", font=("Microsoft YaHei UI", 10)).grid(row=2, column=0, sticky=tk.W,
                                                                                       pady=5)
        tk.Entry(frame, textvariable=self.target_var, font=("Microsoft YaHei UI", 10)).grid(row=2, column=1, sticky=tk.EW, padx=10,
                                                                             pady=5, ipady=3)
        tk.Button(frame, text="浏览...", command=self.select_target, width=10).grid(row=2, column=2, pady=5)

        # 4. 检索关键字
        tk.Label(frame, text="4. 检索关键字:", font=("Microsoft YaHei UI", 10)).grid(row=3, column=0, sticky=tk.NW,
                                                                                     pady=(12, 5))
        tk.Entry(frame, textvariable=self.keyword_var, font=("Microsoft YaHei UI", 10)).grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=10,
                                                            pady=(12, 5), ipady=3)
        tk.Label(frame, text="* 多个关键字请用逗号分隔（支持中英文逗号），支持模糊匹配。",
                 font=("Microsoft YaHei UI", 9), fg="#6C757D").grid(row=4, column=1, sticky=tk.W, padx=10)

        # 5. 控制按钮区 (采用统一样式的着色按钮)
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
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

        # 6. 富文本日志区
        self.log_area = scrolledtext.ScrolledText(
            frame, font=("Microsoft YaHei UI", 10), bg="#F8F9FA", fg="#2C3E50",
            relief=tk.FLAT, padx=12, pady=12, state=tk.DISABLED
        )
        self.log_area.grid(row=6, column=0, columnspan=3, sticky=tk.NSEW, pady=5)

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

        # 使用 self.parent.after 代替原有的 self.root.after
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

        threading.Thread(target=self.process_data, args=(excel_path, source_dir, target_dir, keywords_list),
                         daemon=True).start()

    # ================= 核心算法 =================
    def _clean_filename(self, text):
        return re.sub(r'[\\/:*?"<>|\s]', '_', str(text).strip())

    def _clean_feature_string(self, text):
        if not text: return ""
        return re.sub(r'[^\w\u4e00-\u9fa5]', '', str(text).lower())

    def _is_file_matched(self, combined_feature, emp_name, emp_id, keywords_list):
        clean_feature = self._clean_feature_string(combined_feature)
        clean_name = self._clean_feature_string(emp_name)
        clean_id = self._clean_feature_string(emp_id)

        has_identity = False
        if clean_id and clean_id in clean_feature:
            has_identity = True
        elif clean_name and clean_name in clean_feature:
            has_identity = True

        if not has_identity: return False

        for kw in keywords_list:
            if fuzz.partial_ratio(kw, clean_feature) >= 80:
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

    # ================= 主工作流 =================
    def process_data(self, excel_path, source_dir, target_dir, keywords_list):
        self.log_to_ui(f"🚀 启动自动化处理进程... (当前生效关键字: {len(keywords_list)}个)", "INFO")

        try:
            if not os.path.exists(target_dir): os.makedirs(target_dir)

            self.log_to_ui("📂 正在扫描源目录并构建上下文特征库...", "INFO")
            file_index = []
            for root_dir, _, files in os.walk(source_dir):
                if self._stop_event.is_set():
                    self.log_to_ui("⚠️ 扫描阶段被取消！", "ERROR")
                    return
                self._pause_event.wait()

                parent_folder = os.path.basename(root_dir)
                for file_name in files:
                    ext = os.path.splitext(file_name)[1].lower()
                    if ext in VALID_EXTENSIONS:
                        combined_feature = f"{parent_folder}_{file_name}"
                        file_index.append((combined_feature, file_name, os.path.join(root_dir, file_name)))

            self.log_to_ui(f"✅ 索引构建完成！共检索到 {len(file_index)} 个有效文件。", "SUCCESS")

            df = pd.read_excel(excel_path, dtype=str)
            if "处理状态" not in df.columns: df["处理状态"] = ""
            if "落盘路径" not in df.columns: df["落盘路径"] = ""

            success_count = 0
            for index, row in df.iterrows():
                if self._stop_event.is_set():
                    self.log_to_ui("⚠️ 任务已被强制取消，正在保存已处理进度...", "ERROR")
                    break
                self._pause_event.wait()

                try:
                    emp_id = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                    emp_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                except IndexError:
                    self.log_to_ui("❌ Excel 格式错误：请确保 B 列为工号，C 列为姓名。", "ERROR")
                    break

                if not emp_id and not emp_name: continue

                matched_files = []
                for combined_feature, original_filename, fpath in file_index:
                    if self._is_file_matched(combined_feature, emp_name, emp_id, keywords_list):
                        matched_files.append((original_filename, fpath))

                if not matched_files:
                    df.at[index, "处理状态"] = "⚠️ 未匹配到资质材料"
                    continue

                folder_name = f"{index + 1}_{self._clean_filename(emp_id)}_{self._clean_filename(emp_name)}"
                emp_target_dir = os.path.join(target_dir, folder_name)

                if not os.path.exists(emp_target_dir): os.makedirs(emp_target_dir)

                copied = 0
                seen_hashes = set()

                for fname, fpath in matched_files:
                    if self._stop_event.is_set(): break
                    self._pause_event.wait()

                    file_md5 = self._get_file_md5(fpath)
                    if file_md5:
                        if file_md5 in seen_hashes:
                            self.log_to_ui(f"🔂 拦截到内容重复的文件，已跳过: {fname}", "WARNING")
                            continue
                        seen_hashes.add(file_md5)

                    dst_path = os.path.join(emp_target_dir, fname)
                    base, ext = os.path.splitext(fname)
                    counter = 1
                    while os.path.exists(dst_path):
                        dst_path = os.path.join(emp_target_dir, f"{base}({counter}){ext}")
                        counter += 1

                    shutil.copy2(fpath, dst_path)
                    copied += 1

                df.at[index, "处理状态"] = f"✅ 已成功归档 {copied} 个文件"
                df.at[index, "落盘路径"] = emp_target_dir
                success_count += 1
                self.log_to_ui(f"✅ {emp_name} 最终匹配并保留 {copied} 个独立文件，已归档。", "SUCCESS",
                               hyperlink_path=emp_target_dir)

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

            self.parent.after(0, reset_ui)  # 使用 self.parent.after