# -*- mode: python -*-

block_cipher = None


a = Analysis(['root.py'],
             pathex=['E:\\haha\\receiving-registor'],
             binaries=[],
             datas=[('sdgs.png', '.'), ('sdgs.ico', '.'), ('folder.png', '.'), ('excel.png', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='receiving-registor',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='banana.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='receiving-registor')
