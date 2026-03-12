import os
import json
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD

# ================= 导入各个超级工具模块 =================
from tools.resume_engine import ResumeGeneratorUI
from tools.file_matcher_engine import AutoFileProcessorUI
from tools.word2pdf_engine import WordToPdfUI
from tools.word2img_engine import WordToImageUI
from tools.pdf2img_engine import PdfToImageUI
from tools.img_compress_engine import ImageCompressUI
from tools.word_split_engine import WordSplitUI
from tools.resume_extract_engine import ResumeExtractUI

# 布局配置文件名
LAYOUT_CONFIG_FILE = "app_layout_config.json"


class DraggableNotebook(ttk.Notebook):
    """支持拖拽排序和状态记忆的自定义高级标签页组件"""

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
            # 探测鼠标当前移动到了哪个标签的上方
            target_index = self.index(f"@{event.x},{event.y}")
            if self._dragging_tab_index != target_index:
                # 获取当前被拖拽的真实面板对象
                dragged_tab = self.tabs()[self._dragging_tab_index]
                # 将面板插入到目标位置
                self.insert(target_index, dragged_tab)
                # 更新当前正在拖拽的索引
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
            # 获取当前所有标签的文字名称作为唯一 ID
            tab_order = [self.tab(tab_id, "text") for tab_id in self.tabs()]
            with open(LAYOUT_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({"tab_order": tab_order}, f, ensure_ascii=False)
        except Exception:
            pass

    def restore_layout(self):
        """软件启动时，尝试恢复上一次的标签页顺序"""
        if os.path.exists(LAYOUT_CONFIG_FILE):
            try:
                with open(LAYOUT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    saved_order = config.get("tab_order", [])

                current_tabs = list(self.tabs())
                # 建立 { "📝 智能简历引擎": tab_id } 的映射字典
                text_to_id = {self.tab(tab_id, "text"): tab_id for tab_id in current_tabs}

                # 按照保存的顺序，依次将原本的面板抽出来插到对应位置
                for i, text in enumerate(saved_order):
                    if text in text_to_id:
                        self.insert(i, text_to_id[text])
            except Exception:
                pass


class MainToolbox:
    def __init__(self, root):
        self.root = root
        self.root.title("效能工具箱 V5.0 (支持标签页拖拽定制)")

        # 根据屏幕大小自适应窗口尺寸
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        w, h = int(ws * 0.40), int(hs * 0.85)
        x, y = int((ws / 2) - (w / 2)), int((hs / 2) - (h / 2))
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(950, 700)

        self.setup_global_style()
        self.setup_main_layout()

    def setup_global_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        # 全局字体与颜色设定
        style.configure('.', font=("Microsoft YaHei UI", 10), background="#F0F2F5")

        # 精心调优标签页的视觉风格
        style.configure('TNotebook', background="#E1E5EA", borderwidth=0)

        # 【核心修改 1：让未选中的页签变矮】
        # padding=[左右间距, 上下间距]。将上下间距压低到 4，让后台页签显得扁平、低调
        style.configure('TNotebook.Tab',
                        font=("Microsoft YaHei UI", 10, "bold"),
                        padding=[15, 4],
                        background="#D0D4D9",
                        foreground="#555555",
                        borderwidth=0)

        # 【核心修改 2：让选中的页签动态拔高】
        # expand=[左, 上, 右, 下]。当状态为 'selected' 时，顶部向上扩张 8 个像素！
        style.map('TNotebook.Tab',
                  background=[('selected', '#FFFFFF')],
                  foreground=[('selected', '#0078D7')],
                  expand=[('selected', [2, 8, 2, 0])])

    def setup_main_layout(self):
        # 顶部标题栏
        #header_frame = tk.Frame(self.root, bg="#0078D7", height=60)
        # header_frame.pack(fill=tk.X, side=tk.TOP)
        #  tk.Label(header_frame, text="⚡ 实施交付综合效能工作台",
        #   font=("Microsoft YaHei UI", 16, "bold"), fg="white", bg="#0078D7").pack(side=tk.LEFT, padx=20, pady=15)

        # 提示语
        #tk.Label(header_frame, text="💡 提示: 用鼠标长按上方标签页可自由拖拽排序哦",
        #   font=("Microsoft YaHei UI", 9), fg="#A9D0F5", bg="#0078D7").pack(side=tk.RIGHT, padx=20, pady=15)

        # ==========================================
        # 挂载中心 (使用全新的 DraggableNotebook)
        # ==========================================
        self.notebook = DraggableNotebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. 智能简历引擎
        self.tab1 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab1, text="📝 智能简历引擎")
        ResumeGeneratorUI(self.tab1)

        # 2. Excel数据驱动逆向提取神器
        self.tab8 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab8, text="🧲 简历提取")
        ResumeExtractUI(self.tab8)

        # 3. 文件智能分发系统
        self.tab2 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab2, text="🗂️ 证件资料搜集")
        AutoFileProcessorUI(self.tab2)

        # 4. Word转图片神器
        self.tab4 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab4, text="🖼️ Word转图片")
        WordToImageUI(self.tab4)

        # 5. PDF转图片神器
        self.tab5 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab5, text="📸 PDF转图片")
        PdfToImageUI(self.tab5)

        # 6. 图片智能压缩
        self.tab6 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab6, text="🗜️ 图片智能压缩")
        ImageCompressUI(self.tab6)

        # 7. Word转PDF神器
        self.tab3 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab3, text="🔄 Word转PDF")
        WordToPdfUI(self.tab3)

        # 8. Word按大纲拆分神器
        self.tab7 = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab7, text="✂️ Word拆分神器")
        WordSplitUI(self.tab7)

        # 【核心魔法】：所有标签加载完毕后，读取本地记忆的布局配置，并重新洗牌
        self.notebook.restore_layout()

if __name__ == "__main__":
    import win32event
    import win32api
    import winerror
    from tkinter import messagebox

    # 1. 创建系统级互斥锁 (名称可以自定义，只要独一无二即可)
    mutex_name = "Global\\EfficiencyToolbox_ToughMinded_V5"
    mutex = win32event.CreateMutex(None, False, mutex_name)

    # 2. 探针检测：如果锁已经被占用，说明软件已经在运行了
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        # 创建一个隐藏的 Tk 窗口用来弹窗
        temp_root = TkinterDnD.Tk()
        temp_root.withdraw()
        # 置顶弹窗提示用户
        temp_root.attributes('-topmost', True)
        messagebox.showwarning("提示", "效能工具箱已经在运行中，请勿重复打开！\n请在电脑底部任务栏查找已打开的窗口。")
        temp_root.destroy()
    else:
        # 3. 锁未被占用，正常启动主程序
        root = TkinterDnD.Tk()
        app = MainToolbox(root)
        root.mainloop()