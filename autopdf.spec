# -*- mode: python ; coding: utf-8 -*-

import sys

block_cipher = None

a = Analysis(
    ['src/main.py'],  
    pathex=['c:\\Users\\sa\\Desktop\\AutoPDF_Generator'],  # 添加项目根目录绝对路径
    binaries=[],
    datas=[
        ('templates/*', 'templates'),  # 递归包含模板目录
        ('config.json', '.'),
        # 添加所有静态资源路径
        ('src/*.py', 'src'),
        ('hooks/*.py', 'hooks')
    ],
    hiddenimports=[
        # 明确添加 tkcalendar 核心依赖
        'tkinterweb',
        'babel',
        'babel.dates',
        'pytz',
        
        # 扩展 tkcalendar 所有子模块
        'tkcalendar.*',
        'tkcalendar.calendar_',
        'tkcalendar.popup',
        'tkcalendar.tooltip',
        
        # 补全 docx2pdf 依赖
        'comtypes',
        'comtypes.gen',
        'comtypes.client',
        'win32com',
        'win32timezone',
        'docx2pdf',
        'importlib.metadata',
        
        # 常见缺失依赖补丁
        'PIL._tkinter_finder',
        'pkg_resources.py2_warn'
    ],
    hookspath=['hooks/'],  # 指定自定义钩子目录
    # 启用深层扫描模式
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    # 开启高级日志跟踪
    debug=['all'],
    # 大文件模式优化
    optimize=2
)

# Windows 系统专属配置项
if sys.platform == 'win32':
    from PyInstaller.utils.hooks import collect_submodules
    # 添加系统字体
    a.datas += [
        ('msyh.ttc', 'C:/Windows/Fonts/msyh.ttc', 'DATA')
    ]
    # 确保包含所有必要的系统DLL
    a.binaries += [
        # 添加常见的系统DLL
        ('msvcp140.dll', 'C:\\Windows\\System32\\msvcp140.dll', 'BINARY'),
        ('vcruntime140.dll', 'C:\\Windows\\System32\\vcruntime140.dll', 'BINARY')
    ]

# 生成 plist 文件（macOS 需要）
plist = dict(
    CFBundleName='AutoPDF_Generator',
    CFBundleVersion='1.0.0',
    CFBundleIdentifier='com.example.autopdf',
    NSHumanReadableCopyright='Copyright © 2023 Your Company'
)

# 调整打包策略
pyz = PYZ(a.pure, a.zipped_data)

# 添加 manifest 文件（Windows 需要）
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='AutoPDF_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=True,  # 调试阶段使用控制台模式便于查看错误
    icon=None,  # 暂时不使用图标
    version=None  # 暂时不使用版本信息文件
)
