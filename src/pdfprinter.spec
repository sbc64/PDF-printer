# -*- mode: python -*-

block_cipher = None


a = Analysis(['pdfprinter.py'],
             pathex=['C:\\Users\\sebastianb\\Desktop\\Repositories\\pdfprinter\\src'],
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

a.datas += [('gsdll64.dll',r'C:\Program Files\gs\gs9.19\bin\gsdll64.dll','DATA')]
a.datas += [('gsdll64.lib',r'C:\Program Files\gs\gs9.19\bin\gsdll64.lib','DATA')]
a.datas += [('gswin64c.exe',r'C:\Program Files\gs\gs9.19\bin\gswin64c.exe','DATA')]
a.datas += [('gswin64.exe',r'C:\Program Files\gs\gs9.19\bin\gswin64.exe','DATA')]
a.datas += [('emblem_print.ico','emblem_print.ico','DATA')]

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='pdfprinter',
          debug=False,
          strip=False,
          upx=True,
          console=False,
 					icon='emblem_print.ico')
