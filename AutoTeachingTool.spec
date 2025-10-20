# -*- mode: python ; coding: utf-8 -*-
import os
import customtkinter
from PyInstaller.building.datastruct import normalize_toc

# --- CẤU HÌNH POPPLER (CHÚNG TA KHÔNG DÙNG a.binaries NỮA) ---
# Chỉ cần thiết lập biến cho code PyInstaller biết các thư viện Python cần thiết
# ----------------------------------------------------


a = Analysis(
    ['app_main_v3A.py'],
    pathex=['.'], # Thêm thư mục hiện tại vào PATH
    binaries=[], 
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

# Lấy đường dẫn tới thư mục cài đặt CustomTkinter
customtkinter_dir = os.path.dirname(customtkinter.__file__)

# 1. THÊM DATAS (CustomTkinter Assets và index_backup.html)
# Định dạng: (tên_nguồn, tên_đích, 'DATA')
a.datas += [
    # CustomTkinter assets
    (os.path.join(customtkinter_dir, 'assets'), 'customtkinter/assets', 'DATA'),
    (os.path.join(customtkinter_dir, 'tcl_assets'), 'customtkinter/tcl_assets', 'DATA'),
    
    # File HTML Template (cho Tab 4)
    ('index_backup.html', '.', 'DATA'), 
]

# 2. KHÔNG THÊM BINARIES CHO POPPLER

pyz = PYZ(a.pure)

# Cấu hình EXE cho GUI (--windowed)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AutoTeachingTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False, # <-- Đảm bảo GUI chạy đúng
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# KHỐI COLLECT: Dùng chế độ --onedir để tạo thư mục phân phối (dễ copy Poppler vào)
coll = COLLECT(
    exe,
    normalize_toc(a.binaries), 
    normalize_toc(a.datas),     
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AutoTeachingTool', # Tên thư mục phân phối
)