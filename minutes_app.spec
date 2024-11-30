# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['minutes_app.py'],
    pathex=[],
    binaries=[
        ('ffmpeg.exe', '.'),
        ('ffprobe.exe', '.'),
    ],
    datas=[
        ('template.docx', '.'),
        ('settings.json', '.'),
        # フォントやリソースファイルがある場合はここに追加
    ],
    hiddenimports=[
        'tkinter',
        'google.generativeai',
        'openpyxl',
        'dotenv',
        'docx',
        'PIL',
        'PIL._tkinter_finder',
        'win32gui',
        'win32con',
        'win32api',
        'numpy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='minutes_app',
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
    uac_admin=True,
    version='file_version_info.txt',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='minutes_app',
)
