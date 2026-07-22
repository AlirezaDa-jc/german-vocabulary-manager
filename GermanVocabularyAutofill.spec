# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_dir = Path.cwd()

datas = [
    (str(project_dir / "core"), "core"),
    (str(project_dir / "config.py"), "."),
    (str(project_dir / "create_excel.py"), "."),
    (str(project_dir / "autofill.py"), "."),
]

if (project_dir / "vocabulary.xlsx").exists():
    datas.append((str(project_dir / "vocabulary.xlsx"), "."))

if (project_dir / "data").exists():
    datas.append((str(project_dir / "data"), "data"))

if (project_dir / "audio").exists():
    datas.append((str(project_dir / "audio"), "audio"))

if (project_dir / "images").exists():
    datas.append((str(project_dir / "images"), "images"))

a = Analysis(
    ["app.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "openpyxl",
        "requests",
        "deep_translator",
        "deep_translator.google",
        "tkinter",
    ],
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
    [],
    exclude_binaries=True,
    name="GermanVocabularyManager",
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

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="GermanVocabularyManager",
)

app = BUNDLE(
    coll,
    name="GermanVocabularyManager.app",
    icon="icon.icns" if (project_dir / "icon.icns").exists() else None,
    bundle_identifier="com.alirezada.germanvocabularymanager"
)