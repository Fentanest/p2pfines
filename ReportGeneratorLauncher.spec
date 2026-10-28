# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path(SPECPATH).resolve()

datas = [
    (str(project_root / "templates"), "templates"),
    (str(project_root / "apps_script"), "apps_script"),
]
for workbook_name in ("june.xlsx", "업무보고서 작성기_v5.4_통합.xlsx"):
    workbook_path = project_root / workbook_name
    if workbook_path.exists():
        datas.append((str(workbook_path), "."))

hiddenimports = [
    "google_auth_httplib2",
    "google.oauth2.credentials",
    "google.oauth2.service_account",
    "googleapiclient.discovery",
    "msoffcrypto",
    "pythoncom",
    "pywintypes",
    "win32com",
    "win32com.client",
    "xlrd",
]

a = Analysis(
    [str(project_root / "desktop_launcher.py")],
    pathex=[str(project_root)],
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
    [],
    exclude_binaries=True,
    name="ReportGeneratorLauncher",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name="ReportGeneratorLauncher",
)
