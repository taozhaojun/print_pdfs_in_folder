import sys
from cx_Freeze import setup, Executable

# Include any additional packages or modules
build_exe_options = {
    "packages": ["os", "win32print", "win32api", "tkinter"],
    "includes": ["tkinter"],
    "include_files": [],
    "excludes": []
}

# Base setup for GUI applications on Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # This hides the console window

setup(
    name="PDF Printer",
    version="1.0",
    description="A simple PDF printing application.",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon="printer.ico")]
)