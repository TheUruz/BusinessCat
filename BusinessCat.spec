# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['components\\main.py'],
             pathex=['C:\\Users\\Fabio.Magrotti.CSI\\source\\repos\\BusinessCat'],
             binaries=[],
             datas=[('C:\\Users\\Fabio.Magrotti.CSI\\source\\repos\\BusinessCat\\venv\\Lib\\site-packages\\google_api_python_client-1.12.8.dist-info\\*', 'google_api_python_client-1.12.18.dist-info')],
             hiddenimports=[],
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
          [],
          exclude_binaries=True,
          name='BusinessCat',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='config_files\\imgs\\Cat.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='BusinessCat')
