# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec 文件 - 用于 GitHub Actions 构建 Windows 可执行文件
解决 "failed to load python dll" 问题：显式打包 VC++ 运行时 DLL
"""

import os
import sys
import glob

block_cipher = None

# ---- 自动查找 VC++ 运行时 DLL ----
vc_runtime_binaries = []

# 从 Python 安装目录查找
python_dir = os.path.dirname(sys.executable)
for dll_name in ['vcruntime140.dll', 'vcruntime140_1.dll', 'msvcp140.dll',
                 'concrt140.dll', 'vccorlib140.dll', 'ucrtbase.dll']:
    for search_dir in [python_dir, os.path.join(python_dir, 'DLLs'),
                       r'C:\Windows\System32']:
        dll_path = os.path.join(search_dir, dll_name)
        if os.path.exists(dll_path):
            vc_runtime_binaries.append((dll_path, '.'))
            print(f"[FOUND] {dll_name} -> {dll_path}")
            break

# 查找 api-ms-win-crt-*.dll (Universal CRT)
for search_dir in [python_dir, os.path.join(python_dir, 'DLLs')]:
    for dll in glob.glob(os.path.join(search_dir, 'api-ms-win-crt-*.dll')):
        vc_runtime_binaries.append((dll, '.'))
        print(f"[FOUND] {os.path.basename(dll)}")
    for dll in glob.glob(os.path.join(search_dir, 'api-ms-win-core-*.dll')):
        vc_runtime_binaries.append((dll, '.'))

a = Analysis(
    ['generate_with_images.py'],
    pathex=[],
    binaries=vc_runtime_binaries,
    datas=[('data1', 'data1')],
    hiddenimports=[
        'win32com.client',
        'openpyxl',
        'xlrd',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'ctypes',
        'json',
        'csv',
        'copy',
        'shutil',
        'threading',
        'subprocess',
        'collections',
        'datetime',
        'tempfile',
        're',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=True,  # 打包私有程序集，避免依赖系统 DLL
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # onedir 模式
    name='喜报生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 不压缩，避免某些杀毒软件误报
    console=False,  # 无控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='喜报生成器',
)
