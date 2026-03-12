#公共格式化工具
#这个文件专门存放通用的数据处理逻辑，未来你写“工资条引擎”时也能直接调用它，实现彻底的代码复用。
import pandas as pd


def parse_dynamic_list(text_data):
    """
    解析带有换行符和竖线分隔符的动态列表数据，
    将其转换为可供 Jinja2 引擎循环渲染的字典列表。
    """
    result_list = []

    # 如果数据为空或者不是字符串，直接返回空列表
    if not isinstance(text_data, str) or not text_data.strip():
        return result_list

    lines = text_data.split('\n')
    for line in lines:
        if line.strip():
            # 按竖线切分并去除两端空格
            parts = [p.strip() for p in line.split('|')]
            row_dict = {}
            for idx, part in enumerate(parts):
                # 生成类似 {'c1': '...', 'c2': '...'} 的结构
                row_dict[f'c{idx + 1}'] = part
            result_list.append(row_dict)

    return result_list


def filter_date(value, fmt='%Y-%m-%d'):
    """
    日期过滤器：将 Excel 中的时间格式化为指定的字符串格式。
    如果解析失败（如遇到“至今”、“保密”等文本），则原样返回。
    """
    if pd.isna(value) or str(value).strip() == '':
        return ""

    try:
        return pd.to_datetime(value).strftime(fmt)
    except Exception:
        return str(value)


def filter_num(value, fmt='.2f'):
    """
    数字过滤器：将 Excel 中的数字格式化为指定的小数位数。
    如果解析失败（如遇到非数字文本），则原样返回。
    """
    if pd.isna(value) or str(value).strip() == '':
        return ""

    try:
        return format(float(value), fmt)
    except Exception:
        return str(value)


