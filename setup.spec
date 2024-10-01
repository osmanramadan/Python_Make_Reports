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
pdf_hiddenimports = collect_submodules('PyPDF2')
convert_numbers_hiddenimports = collect_submodules('convert_numbers')
docx2pdf_hiddenimports = collect_submodules('docx2pdf')

# Combine all hidden imports
hidden_imports = pyqt6_hiddenimports + pillow_hiddenimports + docx_hiddenimports + pdf_hiddenimports + \
                 convert_numbers_hiddenimports + docx2pdf_hiddenimports

# Specify the folder containing images
images_folder = os.path.join('E:/school_reports_python/python_make_reports/images')  # Update with the correct path
icons_folder = os.path.join('E:/school_reports_python/python_make_reports/icons')
app_db_file = os.path.join('E:/school_reports_python/python_make_reports/app.db') 


# Datas list - specify the images folder and any other required files
datas = [
    (images_folder, 'images'),
    (icons_folder, 'icons'),
    (app_db_file,'app.db')
]

# Define the PyInstaller configuration
a = Analysis(
    [main_file],
    pathex=['.'],  # Optional: add paths where your modules or data files exist
    hiddenimports=hidden_imports,
    binaries=[],
    datas=datas,  # Add the datas (images folder) here
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
    icon='E:/school_reports_python/python_make_reports/icons/icon.ico',
    version='1.0.0',  # Specify your application version
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
