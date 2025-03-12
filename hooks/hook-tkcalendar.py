from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata

# 自动收集所有依赖模块
datas, binaries, hiddenimports = collect_all('tkcalendar', include_py_files=True)

# 包含 Babel 语言包
datas += copy_metadata('babel')
hiddenimports += collect_submodules('babel')
