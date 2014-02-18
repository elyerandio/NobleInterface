# -*- mode: python -*-
a = Analysis(['NobleInterface.py'],
             pathex=['D:\\Clients\\Docomo\\Nobel Interface\\DocomoNobleInterface'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='NobleInterface.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False )
