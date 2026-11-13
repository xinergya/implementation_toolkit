"""通用路径与文件工具函数，供各引擎复用，消除重复的纯函数实现。"""
import os
import re


def strip_quotes(path):
    """去除用户输入路径两端可能存在的引号与空白。"""
    return str(path or "").strip().strip('"').strip("'")


def sanitize_filename(name, max_len=None, replace_char="_"):
    """将字符串清洗为合法的文件名片段。

    - 去除 Windows 非法字符（\\ / : * ? " < > | 及换行制表符），替换为 replace_char；
    - 合并连续的替换字符；
    - 可选按 max_len 截断。
    """
    clean = re.sub(r'[\\/:*?"<>|\r\n\t]', replace_char, str(name))
    clean = re.sub(re.escape(replace_char) + r'+', replace_char, clean).strip()
    if max_len and len(clean) > max_len:
        clean = clean[:max_len]
    return clean


def format_size(size_bytes):
    """将字节数格式化为人类可读字符串（B / KB / MB）。"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def is_valid_word_file(path):
    """校验是否为合法的 Word 文档且非隐藏临时文件（~$ 开头）。"""
    name = os.path.basename(path)
    return name.lower().endswith(('.doc', '.docx')) and not name.startswith('~$')
