# -*- mode: python -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['C:\\Program Files\\NRB Excel Refresh Links'],
             binaries=[],
             datas=[('relink.ico','.')],
             hiddenimports=['sip'],
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
          name='Refresh Excel Links',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='relink.ico')
