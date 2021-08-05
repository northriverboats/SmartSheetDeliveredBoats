# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['smartsheet__delivered_boats.py'],
             pathex=['/home/fwarren/builds/smartsheet__delivered_boats'],
             binaries=[],
             datas=[
                 ('.env','.'),
             ],
             hiddenimports=[
                 'smartsheet.reports',
                 'emailer.emailer'
             ],
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
          name='smartsheet__delivered_boats',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
