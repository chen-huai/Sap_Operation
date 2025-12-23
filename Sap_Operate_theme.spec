# -*- mode: python ; coding: utf-8 -*-
import os
import sys
import locale

# 设置编码环境
if sys.platform.startswith('win'):
    import codecs
    # 设置系统默认编码
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except:
        pass

from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_all

# 基础配置
app_name = 'Sap_Operate_theme'
entry_point = 'Sap_Operate_theme.py'
icon_file = 'Sap_Operate_Logo.ico'

# 收集数据文件和依赖
datas = []
binaries = []
hiddenimports = []

# 收集主要依赖库（优化 pandas 收集）
print("收集主要依赖库...")
for lib in ['openpyxl', 'xlrd']:
    try:
        tmp_ret = collect_all(lib)
        datas += tmp_ret[0]
        binaries += tmp_ret[1]
        hiddenimports += tmp_ret[2]
        print(f"[OK] 收集 {lib} 完成")
    except Exception as e:
        print(f"[FAIL] 收集 {lib} 失败: {e}")

# 特别处理 numpy（修复 docstring TypeError）
print("收集 numpy 核心模块...")
try:
    from PyInstaller.utils.hooks import collect_data_files
    numpy_datas = collect_data_files('numpy')
    datas.extend(numpy_datas)
    print(f"[OK] 收集 numpy 数据文件完成")
except Exception as e:
    print(f"[FAIL] 收集 numpy 数据失败: {e}")

# 添加 numpy 关键隐藏导入（修复运行时 TypeError）
numpy_core_hidden = [
    'numpy',
    'numpy.core',
    'numpy.core._multiarray_umath',      # 关键：修复 docstring 错误
    'numpy.core._multiarray_tests',
    'numpy.core.multiarray',
    'numpy.core.umath',
    'numpy.core.umath_tests',
    'numpy.core.numeric',
    'numpy.core.fromnumeric',
    'numpy.core.shape_base',
    'numpy.core.c_internal',
    'numpy.core._methods',
    'numpy.core._dtype',
    'numpy.core._exceptions',
    'numpy.linalg',
    'numpy.linalg.lapack_lite',
    'numpy.linalg._umath_linalg',
    'numpy.random',
    'numpy.random.mtrand',
    'numpy.fft',
    'numpy.fft._pocketfft_internal',
    'numpy.polynomial',
    'numpy._core',
    'numpy._core._multiarray_umath',
    'numpy.__config__',
    'numpy.core.overrides',               # 关键：docstring 相关
]
hiddenimports.extend(numpy_core_hidden)
print(f"[OK] 添加 numpy 核心隐藏导入完成")


# 收集 numpy 二进制文件（.pyd 文件）
print("收集 numpy 二进制文件...")
try:
    import numpy
    numpy_path = Path(numpy.__file__).parent

    # 收集 numpy.core 目录中的 .pyd 文件
    core_path = numpy_path / 'core'
    if core_path.exists():
        for pyd_file in core_path.glob('*.pyd'):
            if pyd_file.name not in ['_multiarray_umath.pyd']:
                binaries.append((str(pyd_file), 'numpy/core'))
                print(f"  [添加] {pyd_file.name}")

    # 收集 numpy.linalg 目录中的 .pyd 文件
    linalg_path = numpy_path / 'linalg'
    if linalg_path.exists():
        for pyd_file in linalg_path.glob('*.pyd'):
            binaries.append((str(pyd_file), 'numpy/linalg'))
            print(f"  [添加] linalg/{pyd_file.name}")

    # 收集 numpy.fft 目录中的 .pyd 文件
    fft_path = numpy_path / 'fft'
    if fft_path.exists():
        for pyd_file in fft_path.glob('*.pyd'):
            binaries.append((str(pyd_file), 'numpy/fft'))
            print(f"  [添加] fft/{pyd_file.name}")

    print("[OK] numpy 二进制文件收集完成")

except Exception as e:
    print(f"[WARNING] 收集 numpy 二进制文件失败: {e}")

# 特别处理 pandas（使用 collect_all 彻底收集，修复 DLL load failed 问题）
try:
    from PyInstaller.utils.hooks import collect_all

    # 使用 collect_all 收集 pandas 的所有依赖（数据、二进制、子模块）
    # 这是解决 aggregations.pyd DLL 加载失败的关键
    pandas_datas, pandas_binaries, pandas_hiddenimports = collect_all('pandas')

    # 将收集到的内容添加到对应列表
    datas.extend(pandas_datas)
    binaries.extend(pandas_binaries)
    hiddenimports.extend(pandas_hiddenimports)

    print(f"[OK] 收集 pandas 完整依赖完成:")
    print(f"  - 数据文件: {len(pandas_datas)} 个")
    print(f"  - 二进制文件: {len(pandas_binaries)} 个")
    print(f"  - 隐藏导入: {len(pandas_hiddenimports)} 个")

    # 额外确保关键的 window 模块被包含
    window_specific = [
        'pandas._libs.window.aggregations',
        'pandas._libs.window.indexers',
        'pandas.core.window.ewm',
    ]
    for imp in window_specific:
        if imp not in hiddenimports:
            hiddenimports.append(imp)
    print(f"[OK] 添加 pandas window 特定导入完成")

except Exception as e:
    print(f"[FAIL] 使用 collect_all 收集 pandas 失败: {e}")
    print(f"[INFO] 尝试备用方案...")

    # 备用方案：使用 collect_data_files + collect_submodules
    try:
        from PyInstaller.utils.hooks import collect_submodules, collect_data_files

        pandas_window_modules = collect_submodules('pandas._libs.window')
        hiddenimports.extend(pandas_window_modules)
        print(f"[OK] 备用方案: 收集 pandas._libs.window 子模块完成 ({len(pandas_window_modules)} 个)")

        pandas_datas = collect_data_files('pandas')
        datas.extend(pandas_datas)
        print(f"[OK] 备用方案: 收集 pandas 数据文件完成")

        # 手动收集二进制文件
        import pandas
        pandas_libs_path = Path(pandas.__file__).parent / '_libs'

        for root_dir in [pandas_libs_path, pandas_libs_path / 'tslibs', pandas_libs_path / 'window']:
            if root_dir.exists():
                for pyd_file in root_dir.glob('*.pyd'):
                    dest = 'pandas/_libs' + str(root_dir.relative_to(pandas_libs_path)).replace('\\', '/')
                    binaries.append((str(pyd_file), dest))
                    print(f"  [添加] {pyd_file.name}")

    except Exception as e2:
        print(f"[FAIL] 备用方案也失败: {e2}")

# 收集 PyQt5 相关
print("收集 PyQt5 相关...")
pyqt5_hidden = [
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.QtPrintSupport',
    'PyQt5.QtNetwork',
    'PyQt5.QtSvg',
    'PyQt5.QtXml',
    'PyQt5.QtSql',
    'qt_material',
    'qt_material.resources',
]
hiddenimports.extend(pyqt5_hidden)

# 收集 PDF 处理库
print("收集 PDF 处理库...")
pdf_hidden = [
    'pdfplumber',
    'pdfminer.six',
    'pypdfium2',
    'fitz',  # PyMuPDF
]
hiddenimports.extend(pdf_hidden)

# 收集 Windows 特定库
print("收集 Windows 特定库...")
win_hidden = [
    'win32com.client',
    'win32com.shell',
    'pythoncom',
    'pywintypes',
    'win32gui',
    'win32api',
    'win32con',
]
hiddenimports.extend(win_hidden)

# 收集数据处理库
print("收集数据处理库...")
data_hidden = [
    'chinese_calendar',
    'easyocr',
    'packaging',
]
hiddenimports.extend(data_hidden)

# 收集 setuptools 和 jaraco 依赖（修复运行时错误）
print("收集 setuptools 依赖...")
setuptools_hidden = [
    'setuptools',
    'jaraco.text',
    'jaraco.collections',
    'jaraco.functools',
    'jaraco.context',
    'jaraco.classes',
    'autocommand',
    'more_itertools',
    'pkg_resources',
    'importlib_metadata',
    'zipp',
]
hiddenimports.extend(setuptools_hidden)

# 收集自动更新模块
print("收集自动更新模块...")
auto_updater_hidden = [
    'auto_updater',
    'auto_updater.config',
    'auto_updater.github_client',
    'auto_updater.download_manager',
    'auto_updater.backup_manager',
    'auto_updater.update_executor',
    'auto_updater.ui',
    'auto_updater.ui.dialogs',
    'auto_updater.ui.progress_dialog',
    'auto_updater.ui.widgets',
    'auto_updater.ui.ui_manager',
    'auto_updater.ui.resources',
]
hiddenimports.extend(auto_updater_hidden)

# 添加数据文件
print("添加数据文件...")
data_files = [
    ('Sap_Operate_Ui.ui', '.'),
    ('Table_Ui.ui', '.'),
    ('chicon.py', '.'),
    ('chicon.qrc', '.'),
    ('CLAUDE.md', '.'),
    ('pyi_rth_numpy_fix.py', '.'),
]

# 安全添加数据文件，处理中文路径
def safe_add_data_file(file_path, dest='.'):
    """安全添加数据文件，处理编码问题"""
    try:
        # 确保路径使用正确的编码
        if isinstance(file_path, str):
            file_path_bytes = file_path.encode('utf-8')
            file_path = file_path_bytes.decode('utf-8')

        if os.path.exists(file_path):
            datas.append((file_path, dest))
            print(f"[OK] 添加数据文件: {file_path}")
            return True
        else:
            print(f"[FAIL] 数据文件不存在: {file_path}")
            return False
    except Exception as e:
        print(f"[FAIL] 处理数据文件失败 {file_path}: {e}")
        return False

# 检查文件是否存在并添加到 datas
for file_src, dest in data_files:
    safe_add_data_file(file_src, dest)

# 添加图标文件
if os.path.exists(icon_file):
    datas.append((icon_file, '.'))

# 排除不必要的模块
excludes = [
    'matplotlib',
    'scipy',
    'IPython',
    'jedi',
    'parso',
    'pdoc',
    'pydoc_data',
    'pip',
    # 注意：不排除 wheel、setuptools，因为 jaraco.* 是其依赖
    'tests',
    'unittest',
    'doctest',
]

# 分析配置
a = Analysis(
    [entry_point],
    pathex=[os.getcwd()],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyi_rth_numpy_fix.py'],  # numpy docstring 修复
    excludes=excludes,
    noarchive=False,
    optimize=2,  # 优化级别
)

# 去重处理
seen = set()
unique_datas = []
for item in a.datas:
    if isinstance(item, tuple) and len(item) == 2:
        key = (item[0], item[1])
        if key not in seen:
            seen.add(key)
            unique_datas.append(item)
    else:
        unique_datas.append(item)
a.datas = unique_datas

# PYZ 压缩
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# 可执行文件配置
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'python*.dll'],
    runtime_tmpdir=None,
    console=False,  # 无控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[icon_file] if os.path.exists(icon_file) else [],
)

print("[OK] PyInstaller 配置文件生成完成")
