import sys
from cx_Freeze import setup, Executable


build_exe_options = {"packages": ["os","time","webbrowser","datetime","openpyxl","pynput", "subprocess", "win32api"], "excludes": ["tkinter"]}


base = None


setup(  name = "Zoom_joiner",
        version = "0.1",
        description = "Joins Zoom meetings for you!",
        options = {"build_exe": build_exe_options},
        executables = [Executable("Zoom_Joiner.py", base=base)])