# -*- mode: python -*-

block_cipher = None


a = Analysis(['kp_center.py'],
             pathex=['D:\\Programming\\python\\kp_center'],
             binaries=[],
             datas=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='kp_center',
          debug=False,
          strip=False,
          upx=True,
