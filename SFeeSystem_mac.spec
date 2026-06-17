# -*- mode: python ; coding: utf-8 -*-
# SFeeSystem macOS 打包配置
import os

block_cipher = None
project_root = '/Users/cxd/excelbest'

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=[],
    datas=[
        (os.path.join(project_root, 'data', 'icons'), 'data/icons'),
        (os.path.join(project_root, 'data', 'config'), 'data/config'),
    ],
    hiddenimports=[
        'PyQt5.sip',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'sqlalchemy',
        'sqlite3',
        'pandas',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SFeeSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SFeeSystem',
)

app = BUNDLE(
    coll,
    name='SFeeSystem.app',
    icon=os.path.join(project_root, 'data', 'icons', 'dasheng.icns'),
    bundle_identifier='com.dasheng.sfeesystem',
    info_plist={
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
        'NSHumanReadableCopyright': 'Copyright (c) 大圣智慧软件',
        'CFBundleName': 'SFeeSystem',
        'CFBundleDisplayName': '大圣快递账单结算系统',
    },
)
