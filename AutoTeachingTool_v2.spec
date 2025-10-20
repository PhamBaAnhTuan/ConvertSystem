# -*- mode: python ; coding: utf-8 -*-
import os
import customtkinter

# --- CẤU HÌNH POPPLER (THAY THẾ ĐƯỜNG DẪN CỦA BẠN) ---
POPPLER_BIN_DIR = "D:\\MY_CODE\\mhtml2docx\\Release-25.07.0-0\\poppler-25.07.0\\Library\\bin"
# ----------------------------------------------------


a = Analysis(
    ['app_main_v3A.py'],
    pathex=['.'],
    binaries=[], # Khởi tạo rỗng
    datas=[],    # Khởi tạo rỗng
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

# 1. THÊM DATAS (SỬ DỤNG CÚ PHÁP 3 PHẦN TỬ: nguồn, đích, loại_code)
a.datas += [
    # CustomTkinter assets
    (os.path.join(customtkinter_dir, 'assets'), 'customtkinter/assets', 'DATA'),
    (os.path.join(customtkinter_dir, 'tcl_assets'), 'customtkinter/tcl_assets', 'DATA'),
    
    # File HTML Template (cho Tab 4)
    ('index_backup.html', '.', 'DATA'), 
]

# 2. THÊM BINARIES (Poppler)
a.binaries += [(POPPLER_BIN_DIR, 'poppler', 'BINARY')]


pyz = PYZ(a.pure)

# Cấu hình EXE cho GUI (console=False)
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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# ----------------------------------------------------------------------
# KHỐI SỬA LỖI COLLECT: Buộc tất cả binaries và datas có đủ 3 phần tử.
# Chúng ta dùng hàm normalize_toc để khắc phục lỗi 'expected 3, got 2'.
# ----------------------------------------------------------------------

from PyInstaller.building.datastruct import normalize_toc

# Thu thập tất cả các binaries/datas, bao gồm cả các mục PyInstaller tự động thêm
# và đảm bảo chúng có 3 phần tử.

coll = COLLECT(
    exe,
    normalize_toc(a.binaries),  # SỬ DỤNG normalize_toc CHO BINARIES
    normalize_toc(a.datas),     # SỬ DỤNG normalize_toc CHO DATAS
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AutoTeachingTool',
)