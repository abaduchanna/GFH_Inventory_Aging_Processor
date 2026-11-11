# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec (PySide6/Qt build)
Developed by www.3SVerse.com © {_year}
Build command: pyinstaller GFH_Inventory_Aging_Processor_Qt.spec

Pilot spec for the project-wide Tkinter -> PySide6 migration. Ships
alongside GFH_Inventory_Aging_Processor.spec (the Tkinter build) - this
does not replace it. Same business logic, same output files/formats,
same config file locations; only the UI toolkit differs.
"""

block_cipher = None

a = Analysis(
    ['GFH_Inventory_Aging_Processor_Qt.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('GFH_Telecom_Logo.png', '.'),
        ('theme_manager_qt.py', '.'),
        ('header_manager_qt.py', '.'),
        ('native_titlebar_qt.py', '.'),
    ],
    hiddenimports=[
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'shiboken6',
        'openpyxl',
        'requests',
        'theme_manager_qt',
        'header_manager_qt',
        'native_titlebar_qt',
        'PIL',
        'gspread',
        'gspread_formatting',
        'pandas',
        'pyautogui',
        'win32clipboard',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
        'tkinter',
        '_tkinter',
        'tkinterdnd2',
        'pyperclip',
        'torch',
        'torchvision',
        'torchaudio',
        'matplotlib',
        'matplotlib.pyplot',
        'numba',
        'llvmlite',
        'sympy',
        'tensorflow',
        'scipy',
        'sklearn',
        'scikit-learn',
        'gi',
        'pygments',
        'fsspec',
        'tensorboard',
        'IPython',
        'ipython',
        'jupyter',
        'notebook',
        'speech_recognition',
        'SpeechRecognition',
        'imageio',
        'imageio_ffmpeg',
        'soundfile',
        'PyQt6',
        'PyQt5',
        'PySide2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)


splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(60, 316),
    text_size=11,
    text_color='#96a8cd',
    text_default='Starting...',
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GFH_Inventory_Aging_Processor_Qt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='GFH_icon.ico',
)
