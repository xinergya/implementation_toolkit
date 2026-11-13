"""引擎 UI 轻量基类：统一线程事件控制与优雅停机入口。

各引擎继承本基类，在 __init__ 中调用 super().__init__(parent_frame) 即获得
统一初始化的 _stop_event / _pause_event，以及供主窗口关闭时调用的 request_shutdown。

说明：本基类当前只收敛「线程事件控制」这一最通用、风险最低的能力；
各引擎的 update_button_ui / log_to_ui / 树形表格等 UI 方法仍保留在子类中，
后续可在此基础上继续上移复用（详见代码审查报告 P1-4 迁移路径）。
"""
import threading


class BaseEngineUI:
    def __init__(self, parent_frame):
        self.parent = parent_frame
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()

    def request_shutdown(self):
        """静默请求停止后台任务（不弹确认框），供窗口关闭时调用。

        使用 getattr 防御性访问，兼容未声明 _pause_event 的引擎。
        """
        stop = getattr(self, "_stop_event", None)
        if stop is not None:
            stop.set()
        pause = getattr(self, "_pause_event", None)
        if pause is not None:
            pause.set()
