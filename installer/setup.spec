# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['setup.py'],
             pathex=['E:\\excelbest\\installer'],
             binaries=[],
             datas=[('..\\data\\icons\\dasheng.ico', 'data\\icons')],
             hiddenimports=['win32com', 'win32com.client'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='setup',
          icon='..\\data\\icons\\dasheng.ico',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
