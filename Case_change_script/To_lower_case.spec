# -*- mode: python -*-

block_cipher = None


a = Analysis(['Lower_case_anyobject.py'],
             pathex=['U:\\Home1\\zy964c\\My Documents\\Python Scripts\\Case_change_script'],
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
          name='To_lower_case',
          debug=False,
          strip=False,
          upx=True,
          console=True )
