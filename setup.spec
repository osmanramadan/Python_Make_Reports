# Import statements required by PyInstaller
from PyInstaller.utils.hooks import collect_submodules
import os

# Define the main Python file to execute
main_file = "main.py"

# PyQt6 imports (collect all submodules)
pyqt6_hiddenimports = collect_submodules('PyQt6')

# Collect submodules for libraries used in the script
pillow_hiddenimports = collect_submodules('PIL')
docx_hiddenimports = collect_submodules('docx')
convert_numbers_hiddenimports = collect_submodules('convert_numbers')
aspose_words_hiddenimports = collect_submodules('aspose.words')

# Combine all hidden imports
hidden_imports = (
    pyqt6_hiddenimports + 
    pillow_hiddenimports + 
    docx_hiddenimports + 
    convert_numbers_hiddenimports + 
    aspose_words_hiddenimports
)

# Datas list - specify the images folder and any other required files
# We use "." to specify the current working directory, and we tell PyInstaller to place them in the same directory as the .exe
datas = [
    ('images', './images'),  # Moves 'images' folder to the .exe path
    ('icons', './icons'),    # Moves 'icons' folder to the .exe path
    ('app.db', './app.db'),  # Moves 'app.db' to the .exe path
]

# Define the PyInstaller configuration
a = Analysis(
    [main_file],
    pathex=['.'],  # Optional: add paths where your modules or data files exist
    hiddenimports=hidden_imports,
    binaries=[],
    datas=datas,  # Add the datas (images, icons, app.db) here to be placed in the .exe folder
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

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
    upx=True,
    console=False,
    icon='./icons/icon.ico',  # The icon to be used for the .exe
    description='برنامج توثيق للتقارير',  # Description of the application
    company_name='Ersal',  # Your company name
    copyright='© 2024 Ersal',  # Copyright notice
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='برنامج توثيق',
)
