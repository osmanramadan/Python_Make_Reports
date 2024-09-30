# from distutils.core import setup
# import py2exe
# import os

# data_files = [
#     ("icons", ["E:/school_reports_python/python_make_reports/icons/icon.ico"]),
#     ("images", ["E:/school_reports_python/python_make_reports/images"]),
#     ("icons", ["E:/school_reports_python/python_make_reports/icons"]),
#     ("design", ["E:/school_reports_python/python_make_reports/design"]),
#     ("tempPdf", ["E:/school_reports_python/python_make_reports/tempPdf"]),
# ]

# setup(
#     name="YourAppName",
#     version="1.0",
#     description="A brief description of your PyQt6 app",
#     author="Your Name",
#     windows=[{
#         "script": "E:/school_reports_python/python_make_reports/main.py",  # Point to the correct script
#         "icon_resources": [(1, "E:/school_reports_python/python_make_reports/icons/icon.ico")]
#     }],
#     data_files=data_files,
#     package_dir={'': 'src'},  # Tell setuptools to look for packages in `src/`
#     packages=['your_python_package'],
#     options={
#         'py2exe': {
#             'includes': ['PyQt6'],
#             'excludes': ['Tkinter'],
#             'bundle_files': 1,
#             'compressed': True,
#         }
#     },
#     zipfile=None
# )


from setuptools import setup, find_packages

setup(
    name="YourAppName",
    version="1.0",
    description="A simple PyQt6 application",
    packages=find_packages(where='/'),  # Search for packages inside src directory
    package_dir={'': '/'},  # Specify src as the base directory
    windows=[{
        "script": "main.py",  # Your main script inside src
        "icon_resources": [(1, "E:\school_reports_python\python_make_reports\icons\icon.ico")]  # Optional: specify app icon path
    }],
    options={
        'py2exe': {
            'includes': ['PyQt6'],  # Include PyQt6
        }
    }
)
