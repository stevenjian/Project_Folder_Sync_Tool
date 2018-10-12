import gooey
gooey_root = os.path.dirname(gooey.__file__)
#specpath = os.path.dirname(os.path.abspath(SPEC))
gooey_languages = Tree(os.path.join(gooey_root, 'languages'), prefix = 'gooey/languages')
gooey_images = Tree(os.path.join(gooey_root, 'images'), prefix = 'gooey/images')
a = Analysis(['Project_Folder_Sync_Tool.py'],
             pathex=['C://Users//steve//AppData//Local//Programs//Python//Python36//DLLs'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None,
             )
pyz = PYZ(a.pure)

options = [('u', None, 'OPTION'), ('u', None, 'OPTION'), ('u', None, 'OPTION')]

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          options,
          gooey_languages, # Add them in to collected files
          gooey_images, # Same here.
          name='Project_Folder_Sync_Tool',
          debug=False,
          strip=None,
          upx=True,
          console=False,
          windowed=True)
          
#icon=os.path.join(specpath, 'az_icon.ico'))

