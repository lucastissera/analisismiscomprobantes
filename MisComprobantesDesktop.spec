# -*- mode: python ; coding: utf-8 -*-
# PyInstaller: carpeta de salida dist/MisComprobantesAnalisis/ con MisComprobantesAnalisis.exe

block_cipher = None

a = Analysis(
    ["run_desktop.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("templates", "templates"),
        ("auth_users.example.json", "."),
    ],
    hiddenimports=[
        "pandas",
        "openpyxl",
        "flask",
        "jinja2",
        "werkzeug",
        "auth",
        "sumar_imp_total",
        "plantillas_imputacion",
        "i18n",
        "cuit_en_arca",
        "cuit_en_arca.validacion",
        "cuit_en_arca.errores",
        "cuit_en_arca.credenciales",
        "cuit_en_arca.service",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["playwright"],
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
    name="MisComprobantesAnalisis",
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="MisComprobantesAnalisis",
)
