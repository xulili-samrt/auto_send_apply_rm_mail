# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('mail.png', '.'),
        ('config.json', '.'),  # 添加配置文件到打包资源中
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_tmpdir=None,
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
    name='自动原料申请',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx_exclude=[],
    console=False,  # 修改：设置为False以隐藏控制台窗口
    icon='mail.png'  # 可选：设置可执行文件图标
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx_exclude=[],
    upx_flags=[],
    name='自动原料申请',
    icon='mail.png'
)