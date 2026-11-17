# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('docxcompose')

# main_toolbox 用 __import__(字符串) 动态导入引擎模块，PyInstaller 静态分析无法检测，
# 需显式声明 hiddenimports（collect_submodules 在 spec 顶层因 sys.path 未含项目根目录会返回空，
# 故此处用显式列表）。新增引擎时需在此同步补充。
hiddenimports = [
    'pandas', 'openpyxl',
    'tools.resume_engine',
    'tools.resume_extract_engine',
    'tools.file_matcher_engine',
    'tools.word2img_engine',
    'tools.pdf2img_engine',
    'tools.img_compress_engine',
    'tools.word2pdf_engine',
    'tools.word_split_engine',
    'tools.pdf_to_word_engine',
    'tools.batch_img_insert_engine',
    'tools.remove_headers_footers_engine',
    'utils.formatters',
    'utils.constants',
    'utils.pathutil',
    'utils.ui_base',
    'utils.tree_ui_base',
    'utils.image_types',
]


a = Analysis(
    ['main_toolbox.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='main_toolbox',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version.txt',
    icon=['app.ico'],
)
