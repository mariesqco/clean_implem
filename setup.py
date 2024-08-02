# from cx_Freeze import setup, Executable
#
# # Dependencies are automatically detected, but some modules need manual inclusion
# build_exe_options = {
#     "packages": ["os", "tkinter", "pandas", "openpyxl", "re"],  # Add all needed packages here
#     "excludes": ["matplotlib.tests", "numpy.random._examples"],
#     "include_files": []  # Include any files here that your project needs
# }
#
# # GUI applications require a different base on Windows (the default is for a console application).
# base = None
# import sys
# if sys.platform == "win32":
#     base = "Win32GUI"
#
# setup(
#     name="Excel Processor",
#     version="0.1",
#     description="Your application description.",
#     options={"build_exe": build_exe_options},
#     executables=[Executable("main.py", base=base)]
# )

from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but some modules need manual inclusion
build_exe_options = {
    "packages": ["os", "tkinter", "pandas", "openpyxl", "re", "et_xmlfile"],
    "excludes": ["matplotlib.tests", "numpy.random._examples"],
    "include_files": []  # Include any files here that your project needs
}

# GUI applications require a different base on Windows (the default is for a console application).
base = None
import sys
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Excel Processor",
    version="0.1",
    description="Your application description.",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)]
)
