# Import statements required by PyInstaller
from PyInstaller.utils.hooks import collect_submodules
import os

# Define the main Python file to execute
main_file = "main.py"

# PyQt6: Import submodules needed
pyqt6_hiddenimports = collect_submodules('PyQt6.QtWidgets') + \
                      collect_submodules('PyQt6.QtGui') + \
                      collect_submodules('PyQt6.QtCore')

# Pillow (PIL): Only import necessary submodules
pillow_hiddenimports = collect_submodules('PIL.Image') + \
                       collect_submodules('PIL.ImageOps')

# Docx: Import required submodules
docx_hiddenimports = collect_submodules('docx') + \
                     collect_submodules('docx.enum') + \
                     collect_submodules('docx.enum.section') + \
                     collect_submodules('docx.oxml.xmlchemy') + \
                     collect_submodules('docx.oxml.shared') + \
                     collect_submodules('docx.oxml.ns')

# Convert_numbers: Only import needed modules
convert_numbers_hiddenimports = collect_submodules('convert_numbers')

# Pyautogui: Import submodules for automation
pyautogui_hiddenimports = collect_submodules('pyautogui')

# Arabic Reshaper: Import required module
arabic_reshaper_hiddenimports = collect_submodules('arabic_reshaper')

# Bidi: Import required module
bidi_hiddenimports = collect_submodules('bidi')

# ReportLab: Import required modules
reportlab_hiddenimports = collect_submodules('reportlab') + \
                          collect_submodules('reportlab.lib') + \
                          collect_submodules('reportlab.platypus') + \
                          collect_submodules('reportlab.pdfbase') + \
                          collect_submodules('reportlab.pdfbase.ttfonts')

# Combine all hidden imports
hidden_imports = (
    pyqt6_hiddenimports + 
    pillow_hiddenimports + 
    docx_hiddenimports + 
    convert_numbers_hiddenimports + 
    pyautogui_hiddenimports +
    arabic_reshaper_hiddenimports +
    bidi_hiddenimports +
    reportlab_hiddenimports
)

# Exclude unused PyQt6 modules to reduce size
excluded_modules = ['PyQt6.QtMultimedia', 'PyQt6.QtNetwork', 'PyQt6.QtQml']

# Define the PyInstaller configuration
a = Analysis(
    [main_file],
    pathex=['.'],  # Optional: add paths where your modules or data files exist
    hiddenimports=hidden_imports,
    binaries=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=excluded_modules,  # Exclude unnecessary PyQt6 modules
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

# PyZ section
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# EXE section - contains executable settings
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='برنامج توثيق',  # Name of your executable
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Use UPX compression to reduce binary size
    console=False,
    icon='./icons/icon.ico',  # The icon to be used for the .exe
    description='برنامج توثيق للتقارير',  # Description of the application
    company_name='Ersal',  # Your company name
    copyright='© 2024 Ersal',  # Copyright notice
)

# COLLECT section - collects binaries, zipfiles, and datas
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,  # Use UPX compression for additional files
    upx_exclude=[],
    name='برنامج توثيق',
)
