# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all


matplotlib_datas, matplotlib_binaries, matplotlib_hiddenimports = collect_all(
    "matplotlib"
)

a = Analysis(
    ["nutanix_health_check.py"],
    pathex=[],
    binaries=[
        ("vendor/node/node.exe", "runtime/node"),
    ] + matplotlib_binaries,
    datas=[
        ("data", "data"),
        ("vendor/node-runtime/node_modules", "runtime/node_modules"),
    ] + matplotlib_datas,
    hiddenimports=matplotlib_hiddenimports,
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="NutanixHealthCheck",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="Nutanix-Health-Check-1.0.0-Windows-x64",
)
