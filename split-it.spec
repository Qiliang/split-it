# -*- mode: python ; coding: utf-8 -*-
import sys

block_cipher = None

a = Analysis(
    ["app.py"],
    pathex=[SPECPATH],
    binaries=[],
    datas=[
        ("split-it.ico", "."),
        ("paser.py", "."),
        ("main.py", "."),
    ],
    hiddenimports=[
        "wx",
        "wx._core",
        "openai",
        "mammoth",
        "markdownify",
        "markitdown",
        "markitdown.converter_utils",
        "markitdown.converter_utils.docx",
        "markitdown.converter_utils.docx.pre_process",
        "cryptography",
        "cryptography.fernet",
        "cryptography.hazmat",
        "cryptography.hazmat.primitives",
        "cryptography.hazmat.primitives.ciphers",
        "cryptography.hazmat.primitives.kdf",
        "cryptography.hazmat.primitives.kdf.pbkdf2",
        "cryptography.hazmat.backends",
        "cryptography.hazmat.backends.openssl",
        "tqdm",
        "paser",
        "main",
        "PIL",
        "lxml",
        "lxml.etree",
        "zipfile",
        "xml.etree.ElementTree",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "test", "unittest"],
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
    name="split-it",
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
    icon="split-it.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="split-it",
)

# macOS：生成 .app 应用包
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="Split-It.app",
        icon="split-it.ico",
        bundle_identifier="com.splitit.app",
        info_plist={
            "CFBundleShortVersionString": "1.0.0",
            "CFBundleDisplayName": "Split-It",
            "NSHighResolutionCapable": True,
            "LSApplicationCategoryType": "public.app-category.productivity",
        },
    )
