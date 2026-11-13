"""全局常量与错误码映射，供各引擎复用，消除魔法数字。"""

# ============ UI 颜色 ============
COLOR_SUCCESS = "#198754"
COLOR_ERROR = "#DC3545"
COLOR_WARNING = "#D35400"
COLOR_INFO = "#2C3E50"
COLOR_HIGHLIGHT = "#8E44AD"
COLOR_LINK = "#0078D7"
COLOR_MUTED = "#6C757D"
COLOR_PRIMARY_BLUE = "#2980B9"
COLOR_PRIMARY_GREEN = "#1ABC9C"
COLOR_PRIMARY_PURPLE = "#8E44AD"

COLOR_DISABLED_BG = "#E9ECEF"
COLOR_DISABLED_FG = "#ADB5BD"
COLOR_GRAY_BG = "#E0E0E0"
COLOR_GRAY_FG = "#333333"

# ============ COM 错误码（Word 引擎） ============
COM_ERROR_WORD_STUCK = "80080005"          # Word 引擎卡死
COM_ERROR_RPC = "800706BA"                 # RPC 服务器不可用 / Word 崩溃
COM_ERROR_SERVER_FAILED = "-2146959355"    # 服务器运行失败


def com_error_kind(msg):
    """识别底层 COM 异常类型。

    返回 'stuck'（引擎卡死）/ 'crash'（崩溃闪退）/ 'permission'（权限或占用）/ None（未识别）。
    各引擎据此映射到自己的友好文案，避免在多处重复错误码字符串判断。
    """
    if not msg:
        return None
    if COM_ERROR_WORD_STUCK in msg or "服务器运行失败" in msg:
        return "stuck"
    if COM_ERROR_RPC in msg or "RPC 服务器不可用" in msg or "RPC" in msg:
        return "crash"
    if "Permission denied" in msg or "拒绝访问" in msg or "权限" in msg:
        return "permission"
    return None
