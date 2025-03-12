from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

hiddenimports = [
    'comtypes.client',
    'comtypes.gen',
    'win32com',
    'win32timezone',
    'osgeo',
    'importlib.metadata',
    'docx2pdf'
]

datas = collect_data_files('docx2pdf')
# 添加docx2pdf的元数据文件
datas += copy_metadata('docx2pdf')
hiddenimports += collect_submodules('docx2pdf')
