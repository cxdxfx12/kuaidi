# -*- mode: python ; coding: utf-8 -*-
"""
大圣·快递物流派费结算系统 — PyInstaller 打包配置
- onefile 模式：单文件 exe
- windowed 模式：无控制台窗口
- 自动打包图标/默认配置等资源
"""
import os
import sys

# 项目根目录（spec 文件所在目录）
project_root = os.path.dirname(os.path.abspath(SPEC))

block_cipher = None

# ========== datas：需要打包的静态资源 ==========
# 格式 [(源路径, 目标相对路径), ...]
datas = [
    (os.path.join(project_root, 'data', 'icons'), os.path.join('data', 'icons')),
    (os.path.join(project_root, 'data', 'config'), os.path.join('data', 'config')),
]

# ========== binaries：二进制依赖（通常留空，pyinstaller 自动收集）==========
binaries = []

# ========== hiddenimports：PyInstaller 静态分析遗漏的模块 ==========
hiddenimports = [
    # PyQt5 基础组件
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',

    # pandas & 数据处理
    'pandas',
    'openpyxl',
    'xlsxwriter',
    'python_calamine',  # Rust引擎，替代openpyxl读取，快5倍

    # SQLAlchemy + SQLite
    'sqlalchemy',
    'sqlalchemy.dialects.sqlite',

    # 业务模块（显式列出，避免漏包）
    'app',
    'app.core',
    'app.core.fee_calculator',
    'app.core.settlement',
    'app.models',
    'app.models.database',
    'app.models.fee_record',
    'app.models.fee_detail',
    'app.models.station',
    'app.models.courier',
    'app.models.commission_rule',
    'app.models.column_mapping',
    'app.models.user',
    'app.models.path_config',
    'app.services',
    'app.services.calculate_service',
    'app.services.export_service',
    'app.services.column_matcher',
    'app.services.rule_service',
    'app.services.excel_parser',
    'app.ui',
    'app.ui.main_window',
    'app.ui.login_window',
    'app.utils',
]

# ========== hookspath：自定义 hooks 目录（如有）==========
hookspath = [os.path.join(project_root, 'hooks')] if os.path.isdir(os.path.join(project_root, 'hooks')) else []

# ========== runtime_hooks：运行时钩子（在 pyi_rth_pkgres 之前执行，修复兼容性问题）==========
runtime_hooks = [
    os.path.join(project_root, 'rth_fix_pkgres.py'),
]

# ========== excludes：排除不相关的大包（减小体积）==========
excludes = [
    'tkinter',
    'matplotlib',
    'scipy',
    'IPython',
    'notebook',
    'jupyter',
    'pytest',
    'pandas.io.formats.style',
]

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=hookspath,
    runtime_hooks=runtime_hooks,
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 应用图标路径（优先 .ico，不存在则留空）
icon_path = os.path.join(project_root, 'data', 'icons', 'dasheng.ico')
if not os.path.exists(icon_path):
    icon_path = None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='大圣派费结算系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX 不可用，关闭压缩
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,             # GUI 程序，不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
