"""图片类型配置与图片识别处理（智能简历引擎与图片批量插入引擎共用）。

包含：
- 图片类型配置的读写（名称 + 目录 + 识别关键词），持久化到 %APPDATA%；
- 关键词解析与文件名匹配；
- Word 模板占位符扫描；
- 按人员查找图片（兼容「分人子目录」与「平铺」两种结构）；
- 按目录分组的「关键词优先 + 兜底」类型匹配；
- Pillow 图片预处理（EXIF 转正 + 标准格式化）。
"""
import os
import re
import io
import json
from collections import OrderedDict


VALID_IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.bmp', '.jfif', '.gif')

# 默认图片类型示例：keywords 为空表示「兜底类型」，接受该目录下未被其他类型匹配的图片
DEFAULT_IMAGE_TYPES = [
    {"name": "身份证", "dir": "", "keywords": "身份证"},
    {"name": "学历证书", "dir": "", "keywords": "毕业,学位,学士"},
    {"name": "资质证书", "dir": "", "keywords": ""},
    {"name": "社保证明", "dir": "", "keywords": ""},
]


def get_config_path():
    """图片类型配置文件路径（优先 %APPDATA%，回退脚本目录）。"""
    appdata = os.environ.get("APPDATA")
    if appdata:
        return os.path.join(appdata, "EfficiencyToolbox", "img_type_config.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "img_type_config.json")


def load_image_types():
    """加载图片类型配置；文件缺失或解析失败时回退默认类型。兼容旧配置（无 keywords 字段）。"""
    try:
        with open(get_config_path(), 'r', encoding='utf-8') as f:
            data = json.load(f)
        types = data.get("image_types", [])
        if types:
            result = []
            for t in types:
                result.append({
                    "name": str(t.get("name", "")),
                    "dir": str(t.get("dir", "")),
                    "keywords": str(t.get("keywords", "")),
                })
            return result
    except Exception:
        pass
    return [dict(t) for t in DEFAULT_IMAGE_TYPES]


def save_image_types(types):
    """保存图片类型配置（名称 + 目录 + 识别关键词）。"""
    try:
        path = get_config_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({"image_types": types}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存图片类型配置失败：{e}")


def parse_keywords(keywords_str):
    """把逗号分隔的关键词字符串解析为列表（去空、去重）。"""
    seen = []
    for k in (keywords_str or "").replace('，', ',').split(','):
        k = k.strip()
        if k and k not in seen:
            seen.append(k)
    return seen


def match_keywords(img_path, keywords):
    """判断图片文件名是否含任一关键词（不区分大小写）。"""
    fname = os.path.basename(img_path).lower()
    return any(kw.lower() in fname for kw in keywords)


def extract_template_variables(template_path):
    """扫描 Word 模板，提取所有 {{变量}} 占位符变量名。"""
    variables = set()
    if not template_path or not os.path.exists(template_path):
        return variables
    try:
        from docx import Document
        doc = Document(template_path)
        paras = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    paras.extend(cell.paragraphs)
        pattern = re.compile(r'{{\s*([A-Za-z0-9_\u4e00-\u9fa5]+)\s*}}')
        for para in paras:
            for m in pattern.finditer(para.text):
                variables.add(m.group(1))
    except Exception:
        pass
    return variables


# 图片类型占位符的命名关键词（用于从模板占位符中识别图片类型，而非文本字段）
IMAGE_TYPE_KEYWORDS = (
    '图片', '照片', '证书', '证明', '身份', '证件', '头像', '扫描', '资格', '职称',
)


def get_excel_columns(excel_path):
    """读取 Excel 列名（优先「实施人员清单」表，否则第一个表），用于排除文本字段占位符。"""
    if not excel_path or not os.path.exists(excel_path):
        return []
    try:
        import pandas as pd
        xls = pd.ExcelFile(excel_path)
        sheet = '实施人员清单' if '实施人员清单' in xls.sheet_names else 0
        df = pd.read_excel(xls, sheet_name=sheet, nrows=0)
        return [str(c).strip() for c in df.columns]
    except Exception:
        return []


def identify_image_type_variables(variables, excel_columns=None):
    """从模板占位符中识别图片类型名称（排除文本字段）。

    规则：
    1. 排除 Excel 列名（这些是普通文本字段，如姓名/年龄/性别等）；
    2. 剩余占位符中，仅保留名称含图片类型关键词（证书/证明/图片/身份/学历等）的。

    这样可避免把模板中所有 {{表达式}} 都识别为图片类型。
    """
    excel_set = set(excel_columns or [])
    candidates = [v for v in variables if v not in excel_set]
    return [v for v in candidates if any(kw in v for kw in IMAGE_TYPE_KEYWORDS)]


def find_person_images(name, emp_id, search_dir):
    """查找某人员的所有候选图片，兼容「分人子目录」与「平铺」两种结构。

    1. 若 search_dir 下存在以姓名/工号命名的子文件夹，则收集这些子文件夹内的全部图片；
    2. 否则视为平铺目录，收集文件名含姓名/工号的图片。
    """
    images = []
    if not search_dir or not os.path.exists(search_dir):
        return images
    try:
        entries = os.listdir(search_dir)
    except Exception:
        return images

    # 1. 分人子目录模式
    matched_subdirs = []
    for entry in entries:
        full = os.path.join(search_dir, entry)
        if os.path.isdir(full) and (name in entry or (emp_id and emp_id in entry)):
            matched_subdirs.append(full)
    if matched_subdirs:
        for sd in matched_subdirs:
            for root, dirs, files in os.walk(sd):
                for f in files:
                    if f.lower().endswith(VALID_IMAGE_EXTS):
                        images.append(os.path.join(root, f))
        return images

    # 2. 平铺模式
    for f in entries:
        if f.lower().endswith(VALID_IMAGE_EXTS) and (name in f or (emp_id and emp_id in f)):
            images.append(os.path.join(search_dir, f))
    return images


def match_types_for_person(image_types, name, emp_id):
    """对某人员，按目录分组执行「关键词优先 + 兜底」匹配，返回 {类型名: [图片路径]}。"""
    dir_groups = OrderedDict()
    for ti in image_types:
        dir_groups.setdefault(ti["dir"], []).append(ti)

    type_images = {}
    for type_dir, types in dir_groups.items():
        person_imgs = find_person_images(name, emp_id, type_dir)
        used = set()
        # 第一遍：有关键词的类型按文件名匹配
        for ti in types:
            kws = parse_keywords(ti.get("keywords", ""))
            if kws:
                matched = [p for p in person_imgs if p not in used and match_keywords(p, kws)]
                type_images[ti["name"]] = matched
                used.update(matched)
        # 第二遍：关键词为空的类型接受剩余图片（兜底）
        remaining = [p for p in person_imgs if p not in used]
        for ti in types:
            if not parse_keywords(ti.get("keywords", "")):
                type_images[ti["name"]] = remaining
    return type_images


def prepare_image(img_path):
    """用 Pillow 规范化图片为 python-docx 可识别的标准格式，返回 BytesIO。

    解决三类问题：① 某些 JPEG 缺少 JFIF/Exif 头（文件头为 ffd8ffdb），python-docx
    无法识别；② 源图为 WebP/HEIC 等 python-docx 不支持的格式；③ 按 EXIF 方向信息
    转正图片，避免手机拍摄的照片在文档中横向/倒置摆放。转换失败返回 None。
    """
    try:
        from PIL import Image, ImageOps
        with Image.open(img_path) as im:
            fmt = im.format
            im = ImageOps.exif_transpose(im)
            buf = io.BytesIO()
            has_alpha = im.mode in ('RGBA', 'LA', 'P')
            if fmt == 'JPEG':
                im.convert('RGB').save(buf, format='JPEG', quality=92)
            elif has_alpha:
                im.convert('RGBA').save(buf, format='PNG')
            else:
                im.convert('RGB').save(buf, format='PNG')
            buf.seek(0)
            return buf
    except Exception:
        return None
