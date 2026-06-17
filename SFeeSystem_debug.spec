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
        'PIL',
        'PIL.Image',
        'PIL.ImageQt',
        'win32com',
        'win32com.client',
        'pythoncom',
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
        'scipy',
        'scipy.linalg',
        'scipy.special',
        'scipy.ndimage',
        'scipy.sparse',
        'scipy.stats',
        'scipy.integrate',
    ],
    hookspath=[],
    runtime_hooks=['rth_fix_pkgres.py'],
    excludes=[
        'tkinter',
        '_tkinter',
        'Tkinter',
        'matplotlib',
        'IPython',
        'pytest',
        'PySide2',
        'PySide6',
        'PyQt4',
        'PyQt6',
        'PyQt5.QtQuick',
        'PyQt5.QtQuickWidgets',
        'PyQt5.QtQml',
        'PyQt5.QtQmlModels',
        'PyQt5.QtWebEngineCore',
        'PyQt5.QtWebEngineWidgets',
        'PyQt5.QtWebSockets',
        'PyQt5.QtDBus',
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
        'unittest',
        'doctest',
        'lib2to3',
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

exclude_keywords = [
    'Qt5Quick', 'Qt5Qml', 'Qt5WebSockets', 'Qt5WebEngine',
    'Qt5DBus', 'Qt5Svg', 'Qt5Sql', 'Qt5Xml',
    'Qt5Multimedia', 'Qt5Test', 'Qt5PrintSupport', 'Qt5OpenGL',
    'Qt53D', 'Qt5Positioning', 'Qt5Sensors', 'Qt5Serial',
    'qt5_qml', 'qml', 'QtQuick',
    'tcl8', 'tk8', '_tkinter', 'tcl/tcl',
    'lib2to3',
    'libEGL', 'libGLES', 'd3dcompiler', 'opengl32sw',
]

def _should_exclude(path, keywords):
    low = str(path).lower().replace('\\', '/')
    for kw in keywords:
        if kw.lower() in low:
            return True
    return False

a.datas = [(dest, src, typ) for dest, src, typ in a.datas
           if not _should_exclude(dest, exclude_keywords)]

a.binaries = [(dest, src, typ) for dest, src, typ in a.binaries
              if not _should_exclude(dest, exclude_keywords)]

a.zipfiles = [(dest, src, typ) for dest, src, typ in a.zipfiles
              if not _should_exclude(dest, exclude_keywords)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SFeeSystem_debug',
    icon='data/icons/dasheng.ico',
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='SFeeSystem_debug',
)