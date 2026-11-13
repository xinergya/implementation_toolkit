import os
import json
import logging
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD


def _get_layout_config_path():
    """返回布局配置文件路径。

    优先存放到 %APPDATA% 下的固定目录，避免打包为 exe 后工作目录不稳定
    导致布局记忆失效或污染目录；无 APPDATA 时回退到脚本所在目录。
    """
    appdata = os.environ.get("APPDATA")
    if appdata:
        config_dir = os.path.join(appdata, "EfficiencyToolbox")
    else:
        config_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(config_dir, "app_layout_config.json")


# 布局配置文件路径
LAYOUT_CONFIG_FILE = _get_layout_config_path()

# 各引擎标签页（标题, 模块名, 类名），顺序即默认显示顺序，同时用于布局记忆
ENGINE_SPECS = [
    ("📝 智能简历引擎", "tools.resume_engine", "ResumeGeneratorUI"),
    ("🧲 简历提取", "tools.resume_extract_engine", "ResumeExtractUI"),
    ("🗂️ 证件资料搜集", "tools.file_matcher_engine", "AutoFileProcessorUI"),
    ("🖼️ Word转图片", "tools.word2img_engine", "WordToImageUI"),
    ("📸 PDF转图片", "tools.pdf2img_engine", "PdfToImageUI"),
    ("🗜️ 图片智能压缩", "tools.img_compress_engine", "ImageCompressUI"),
    ("🔄 Word转PDF", "tools.word2pdf_engine", "WordToPdfUI"),
    ("✂️ Word拆分神器", "tools.word_split_engine", "WordSplitUI"),
    ("🔄 PDF转Word", "tools.pdf_to_word_engine", "PDFToWordUI"),
    ("🔄 图片批量添加生成目录", "tools.batch_img_insert_engine", "BatchImageInsertUI"),
]


class DraggableNotebook(ttk.Notebook):
    """支持拖拽排序和状态记忆的自定义标签页组件"""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)

        # 绑定鼠标拖拽核心事件
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<B1-Motion>", self.on_motion)
        self.bind("<ButtonRelease-1>", self.on_release)

        self._dragging_tab_index = None

    def on_press(self, event):
        """当鼠标左键按下时，记录当前选中的标签页索引"""
        try:
            # 探测鼠标点中了哪个标签
            self._dragging_tab_index = self.index(f"@{event.x},{event.y}")
        except tk.TclError:
            self._dragging_tab_index = None

    def on_motion(self, event):
        """当鼠标拖动时，动态交换标签页位置"""
        if self._dragging_tab_index is None:
            return
        try:
            target_index = self.index(f"@{event.x},{event.y}")
            if self._dragging_tab_index != target_index:
                dragged_tab = self.tabs()[self._dragging_tab_index]
                self.insert(target_index, dragged_tab)
                self._dragging_tab_index = target_index
        except tk.TclError:
            pass  # 鼠标移出了标签区域，忽略

    def on_release(self, event):
        """当松开鼠标时，将最新的顺序保存到本地"""
        if self._dragging_tab_index is not None:
            self._dragging_tab_index = None
            self.save_layout()

    def save_layout(self):
        """将标签页的顺序持久化保存到 JSON 文件"""
        try:
            tab_order = [self.tab(tab_id, "text") for tab_id in self.tabs()]
            config_dir = os.path.dirname(LAYOUT_CONFIG_FILE)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
            with open(LAYOUT_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({"tab_order": tab_order}, f, ensure_ascii=False)
        except Exception as exc:
            logging.warning("保存标签页布局失败：%s", exc)

    def restore_layout(self):
        """软件启动时，尝试恢复上一次的标签页顺序"""
        if not os.path.exists(LAYOUT_CONFIG_FILE):
            return
        try:
            with open(LAYOUT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                saved_order = config.get("tab_order", [])

            current_tabs = list(self.tabs())
            text_to_id = {self.tab(tab_id, "text"): tab_id for tab_id in current_tabs}

            for i, text in enumerate(saved_order):
                if text in text_to_id:
                    self.insert(i, text_to_id[text])
        except Exception as exc:
            logging.warning("恢复标签页布局失败：%s", exc)


class MainToolbox:
    def __init__(self, root):
        self.root = root
        self.root.title("效能工具箱 V5.0 (支持标签页拖拽定制)")

        # 保存各引擎实例引用，供窗口关闭时统一请求停止后台任务
        self.engine_instances = []

        # 根据屏幕大小自适应窗口尺寸
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        w, h = int(ws * 0.40), int(hs * 0.85)
        x, y = int((ws / 2) - (w / 2)), int((hs / 2) - (h / 2))
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(950, 700)

        # 1. 先初始化 UI 全局样式
        self.setup_global_style()

        # 2. 绘制一个临时的"加载过渡屏"
        self.loading_frame = tk.Frame(self.root, bg="#F0F2F5")
        self.loading_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(self.loading_frame, text="🚀 正在唤醒底层效能引擎，请稍候...",
                 font=("Microsoft YaHei UI", 16, "bold"), fg="#0078D7", bg="#F0F2F5").pack(expand=True)

        # 3. 强制系统立刻把过渡窗口弹出来给用户看
        self.root.update()

        # 4. 绑定窗口关闭协议：关闭前静默停止所有引擎后台线程，避免在已销毁对象上继续回调
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # 5. 延迟 100 毫秒后，再去后台悄悄加载重型引擎
        self.root.after(100, self.load_heavy_engines)

    def setup_global_style(self):
        """恢复被遗漏的样式设置方法"""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=("Microsoft YaHei UI", 10), background="#F0F2F5")
        style.configure('TNotebook', background="#E1E5EA", borderwidth=0)
        style.configure('TNotebook.Tab',
                        font=("Microsoft YaHei UI", 10, "bold"),
                        padding=[15, 4],
                        background="#D0D4D9",
                        foreground="#555555",
                        borderwidth=0)
        style.map('TNotebook.Tab',
                  background=[('selected', '#FFFFFF')],
                  foreground=[('selected', '#0078D7')],
                  expand=[('selected', [2, 8, 2, 0])])

    def load_heavy_engines(self):
        """局部懒加载：逐个容错导入，任一引擎依赖缺失仅影响该引擎，不再拖垮整个应用。"""
        engine_classes = {}
        for title, module_name, class_name in ENGINE_SPECS:
            try:
                mod = __import__(module_name, fromlist=[class_name])
                engine_classes[title] = getattr(mod, class_name)
            except Exception:
                logging.exception("引擎加载失败 %s（%s.%s）", title, module_name, class_name)

        # 引擎加载完毕，销毁过渡屏
        self.loading_frame.destroy()

        # 渲染真正的多标签主界面
        self.setup_main_layout(engine_classes)

    def setup_main_layout(self, engine_classes):
        # ==========================================
        # 挂载中心 (使用全新的 DraggableNotebook)
        # ==========================================
        self.notebook = DraggableNotebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for title, _module_name, _class_name in ENGINE_SPECS:
            tab = tk.Frame(self.notebook, bg="#FFFFFF")
            self.notebook.add(tab, text=title)

            engine_cls = engine_classes.get(title)
            if engine_cls is not None:
                instance = engine_cls(tab)
                self.engine_instances.append(instance)
            else:
                tk.Label(tab, text=f"⚠️ 该模块加载失败\n{title}\n请检查依赖是否完整安装",
                         fg="red", bg="#FFFFFF", font=("Microsoft YaHei UI", 12, "bold"),
                         justify=tk.CENTER).pack(expand=True)

        # 所有标签加载完毕后，读取本地记忆的布局配置，并重新洗牌
        self.notebook.restore_layout()

    def on_close(self):
        """窗口关闭时，静默请求所有引擎停止后台任务，再销毁窗口。"""
        for instance in self.engine_instances:
            stop = getattr(instance, "request_shutdown", None)
            if callable(stop):
                try:
                    stop()
                except Exception:
                    pass
        try:
            self.root.destroy()
        except tk.TclError:
            pass


if __name__ == "__main__":
    import win32event
    import win32api
    import winerror
    from tkinter import messagebox

    # 1. 创建系统级互斥锁，用于防止程序重复启动。
    #    不指定命名空间前缀，默认使用会话级 Local 命名空间（当前登录会话内唯一），
    #    避免 Global\ 前缀在标准用户/受限环境下因缺少 SeCreateGlobalPrivilege 权限导致创建失败。
    mutex_name = "EfficiencyToolbox_ToughMinded_V5"
    mutex = None
    already_running = False
    try:
        mutex = win32event.CreateMutex(None, False, mutex_name)
        already_running = (win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS)
    except Exception as exc:
        # 创建互斥锁失败（权限不足、环境异常等）时，放弃防多开保护，仍允许正常启动，
        # 避免因锁机制异常导致软件无法使用。
        logging.warning("防多开互斥锁创建失败，跳过单实例保护：%s", exc)
        already_running = False

    # 2. 探针检测：如果锁已经被占用，说明软件已经在运行了
    if already_running:
        # 创建一个隐藏的 Tk 窗口用来弹窗
        temp_root = TkinterDnD.Tk()
        temp_root.withdraw()
        temp_root.attributes('-topmost', True)
        messagebox.showwarning("提示", "效能工具箱已经在运行中，请勿重复打开！\n请在电脑底部任务栏查找已打开的窗口。")
        temp_root.destroy()
        if mutex is not None:
            win32api.CloseHandle(mutex)
    else:
        # 3. 锁未被占用，正常启动主程序
        root = TkinterDnD.Tk()
        app = MainToolbox(root)
        try:
            root.mainloop()
        finally:
            # 主实例退出后释放互斥锁句柄，避免内核句柄泄漏
            if mutex is not None:
                win32api.CloseHandle(mutex)
