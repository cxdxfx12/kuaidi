# -*- mode: python ; coding: utf-8 -*-
import sys
import os

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app', 'app'),
        # 只打包必要的数据配置文件：1.6GB 的 app.db 不打包（首次运行由程序在用户目录创建）
        ('data/config/fee_rules.json', 'data/config'),
        ('data/config/default_settings.json', 'data/config'),
        ('data/icons', 'data/icons'),
    ],
    hiddenimports=[
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.QtNetwork',
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.tslibs.conversion',
        'pandas._libs.lib',
        'pandas._libs.reshape',
        'sqlalchemy',
        'sqlalchemy.sql.default_comparator',
        'sqlalchemy.dialects.sqlite',
        'sqlalchemy.ext.declarative',
        'sqlalchemy.orm',
        'openpyxl',
        'openpyxl.reader.excel',
        'openpyxl.writer.excel',
        'xlsxwriter',
        'PIL',
        'PIL.Image',
        'PIL.ImageQt',
        # win32com（用于桌面快捷方式和Excel操作）
        'win32com',
        'win32com.client',
        'pythoncom',
        # 其他必要依赖
        'numpy',
        'numpy.core._methods',
        'numpy.lib.format',
        'datetime',
        'pytz',
        'ctypes',
        'ctypes.wintypes',
        'uuid',
        'hashlib',
        'hmac',
        'itertools',
        'functools',
        'collections',
        'collections.abc',
        'json',
        'csv',
        'os',
        'sys',
        'io',
        'shutil',
        'tempfile',
        'pathlib',
        'time',
        're',
        'math',
        'decimal',
        'logging',
    ],
    hookspath=[],
    runtime_hooks=['rth_fix_pkgres.py'],
    excludes=[
        # 不使用的 GUI 框架
        'tkinter',
        '_tkinter',
        'Tkinter',
        # 不使用的大型库
        'matplotlib',
        'scipy',
        'IPython',
        'pytest',
        'PySide2',
        'PySide6',
        'PyQt4',
        'PyQt6',
        # 不使用的 Qt 模块（只用 Core/Gui/Widgets）
        'PyQt5.QtQuick',
        'PyQt5.QtQuickWidgets',
        'PyQt5.QtQml',
        'PyQt5.QtQmlModels',
        'PyQt5.QtWebEngineCore',
        'PyQt5.QtWebEngineWidgets',
        'PyQt5.QtWebSockets',
        'PyQt5.QtDBus',
        'PyQt5.QtNetwork',
        'PyQt5.QtSvg',
        'PyQt5.QtSql',
        'PyQt5.QtXml',
        'PyQt5.QtXmlPatterns',
        'PyQt5.QtMultimedia',
        'PyQt5.QtMultimediaWidgets',
        'PyQt5.Qt3DCore',
        'PyQt5.Qt3DRender',
        'PyQt5.Qt3DInput',
        'PyQt5.Qt3DAnimation',
        'PyQt5.Qt3DExtras',
        'PyQt5.QtTest',
        'PyQt5.QtPrintSupport',
        'PyQt5.QtOpenGL',
        'PyQt5.uic',
        # 不使用的 Python 内置测试工具
        'unittest',
        'doctest',
        'lib2to3',
        # 注意：email/http/xmlrpc 是 Python 标准库，被 sqlalchemy/pandas 间接依赖
        # 不要排除它们！
        # 排除 pkg_resources/setuptools（解决 pyi_rth_pkgres 崩溃）
        # 只排除顶层包，不排除可能被间接依赖的子模块
        'pkg_resources',
        'setuptools',
        'setuptools._vendor',
        'setuptools.extern',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ======= 额外精简：过滤掉打包进来但实际未使用的文件 =======
# 定义需要排除的文件关键字
exclude_keywords = [
    # Qt 未使用模块（.dll / QtQuick 资源文件）
    'Qt5Quick', 'Qt5Qml', 'Qt5WebSockets', 'Qt5WebEngine',
    'Qt5DBus', 'Qt5Svg', 'Qt5Sql', 'Qt5Xml',
    'Qt5Multimedia', 'Qt5Test', 'Qt5PrintSupport', 'Qt5OpenGL',
    'Qt53D', 'Qt5Positioning', 'Qt5Sensors', 'Qt5Serial',
    'qt5_qml', 'qml', 'QtQuick',
    # Tcl/Tk（我们用 PyQt5，不用 Tkinter）
    'tcl8', 'tk8', '_tkinter', 'tcl/tcl',
    # Python 标准库测试 / 语言开发工具
    'lib2to3',
    # OpenGL 相关驱动
    'libEGL', 'libGLES', 'd3dcompiler', 'opengl32sw',
]

def _should_exclude(path, keywords):
    """判断路径是否应该被排除（大小写不敏感）"""
    low = str(path).lower().replace('\\', '/')
    for kw in keywords:
        if kw.lower() in low:
            return True
    return False

# 过滤 datas
a.datas = [(dest, src, typ) for dest, src, typ in a.datas
           if not _should_exclude(dest, exclude_keywords)]

# 过滤 binaries
a.binaries = [(dest, src, typ) for dest, src, typ in a.binaries
              if not _should_exclude(dest, exclude_keywords)]

# 过滤 zipfiles（内部的 .pyz 中附带的 Python 模块）
a.zipfiles = [(dest, src, typ) for dest, src, typ in a.zipfiles
              if not _should_exclude(dest, exclude_keywords)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SFeeSystem',
    icon='data/icons/dasheng.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

# 强制收集 xlsxwriter（用 binaries 而非 datas，确保 PYTHONPATH 能找到）
import os as _os
_xlw_base = r'C:\Program Files\python\lib\site-packages\xlsxwriter'
if _os.path.isdir(_xlw_base):
    for _root, _dirs, _files in _os.walk(_xlw_base):
        for _f in _files:
            if _f.endswith('.py') and not _f.startswith('.'):
                _src = _os.path.join(_root, _f)
                _rel = _os.path.relpath(_src, _xlw_base).replace('\\', '/')
                _dst = 'xlsxwriter/' + _rel
                a.binaries.append((_dst, _src, 'BINARY'))

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='SFeeSystem',
)
