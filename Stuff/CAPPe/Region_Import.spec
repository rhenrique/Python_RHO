# -*- mode: python -*-

block_cipher = None


a = Analysis(['Region_Import.py'],
             pathex=['c:\\Users\\oliveirh\\Documents\\GitHub\\Python_RHO\\Stuff\\CAPPe'],
             binaries=[],
             datas=[],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Region_Import',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
