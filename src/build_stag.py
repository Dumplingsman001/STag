# -*- coding: utf-8 -*-
"""打包 STag 初版（单文件 Windows exe）。"""

import os

import PyInstaller.__main__

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
os.chdir(HERE)

LOCAL_APP = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "STag"
)
WORK_PATH = os.path.join(LOCAL_APP, "pyinstaller_stag")
DIST_PATH = ROOT
SCRIPT = os.path.join(HERE, "stag.py")
ICON = os.path.join(HERE, "assets", "icon.ico")
LOGO = os.path.join(HERE, "assets", "stag_logo.png")

# 输出到 STag 根目录，保留已有的 config.json
os.makedirs(DIST_PATH, exist_ok=True)

args = [
    SCRIPT,
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", "STag",
    "--icon", ICON,
    "--add-data", f"{ICON};assets",
    "--add-data", f"{LOGO};assets",
    "--hidden-import", "pystray._win32",
    "--collect-submodules", "pystray",
    "--collect-submodules", "winrt",
    "--hidden-import", "openpyxl",
    "--hidden-import", "et_xmlfile",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pythoncom",
    "--collect-submodules", "win32com",
    "--exclude-module", "PySide6",
    "--exclude-module", "PyQt5",
    "--exclude-module", "PyQt6",
    "--exclude-module", "numpy",
    "--exclude-module", "matplotlib",
    "--exclude-module", "torch",
    "--exclude-module", "IPython",
    "--exclude-module", "notebook",
    "--workpath", WORK_PATH,
    "--distpath", DIST_PATH,
    "--specpath", WORK_PATH,
]

PyInstaller.__main__.run(args)
exe_path = os.path.join(DIST_PATH, "STag.exe")
print("BUILD OK:", exe_path, os.path.getsize(exe_path) // (1024 * 1024), "MB")
