# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for LINE Auto RPA
打包成單一執行檔 (--onefile)
"""

import sys
from pathlib import Path

# 定義專案根目錄
spec_root = Path('.').absolute()

# 收集資料檔案
datas = [
    ('config.example.yaml', '.'),
    ('.env.example', '.'),
    ('README.md', '.'),
]

# 收集資料夾（修正版 - 收集資料夾內的所有檔案）
def collect_folder_files(folder_name):
    """收集指定資料夾內的所有檔案"""
    folder = spec_root / folder_name
    files = []
    if folder.exists() and folder.is_dir():
        for file_path in folder.rglob('*'):
            if file_path.is_file():
                # 計算相對路徑
                rel_path = file_path.relative_to(spec_root)
                dest_dir = str(rel_path.parent)
                files.append((str(file_path), dest_dir))
    return files

# 添加必要的資料夾
for folder in ['templates', 'tasks', 'data', 'lists']:
    datas.extend(collect_folder_files(folder))

# 隱藏導入（有些模組需要明確指定）
hiddenimports = [
    'PIL._tkinter_finder',
    'numpy',
    'cv2',
    'pyautogui',
    'pywinauto',
    'pytesseract',
    'apscheduler',
    'dotenv',
    'yaml',
    'loguru',
]

a = Analysis(
    ['main_ui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='LINE_Auto_RPA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 設為 False 以隱藏控制台視窗（GUI 應用）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # 使用自訂圖示
)
