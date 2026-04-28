# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller build configuration for DocFormatter
"""

import sys
import os
from pathlib import Path

block_cipher = None

# Source directory
src_dir = Path("docformatter")

# Main entry point
main_script = src_dir / "gui" / "main_window.py"

a = Analysis(
    [str(main_script)],
    pathex=[],
    binaries=[],
    datas=[
        # Include templates directory if exists
        ("templates", "templates"),
    ],
    hiddenimports=[
        "docformatter",
        "docformatter.core",
        "docformatter.core.formatter",
        "docformatter.core.table_handler",
        "docformatter.core.word2md",
        "docformatter.core.style_mapper",
        "docformatter.core.numbering",
        "docformatter.core.cover_replacer",
        "docformatter.core.header_footer",
        "docformatter.gui",
        "docformatter.gui.main_window",
        "docformatter.gui.template_config",
        "docformatter.gui.batch_process",
        "docformatter.gui.word2md_tab",
        "docformatter.models",
        "docformatter.models.template_config",
        "docformatter.templates",
        "docformatter.templates.template_io",
        "docformatter.utils",
        "docformatter.utils.font_formatter",
        "docformatter.utils.file_utils",
        "docformatter.utils.logger",
        "lxml",
        "openpyxl",
    ],
    hookspath=[],
    hooksconfig={},
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocFormatter",
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
    icon="icon.ico",
    version="version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="DocFormatter",
)

# Windows-specific options
if sys.platform == "win32":
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="DocFormatter",
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
        icon="icon.ico",
        version="version_info.txt",
    )
