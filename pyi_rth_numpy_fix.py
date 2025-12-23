# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller Runtime Hook: numpy docstring fix
修复 numpy.core.overrides 模块中的 docstring TypeError 问题

使用方法：将此文件放置在项目根目录，PyInstaller 会自动识别为 runtime hook
"""

import sys
import builtins

# 保存原始的 add_docstring 函数
_original_add_docstring = None

def safe_add_docstring(obj, docstring):
    """
    安全版本的 add_docstring，处理 docstring 不是字符串的情况

    在 PyInstaller 打包的冻结环境中，C 扩展模块的 __doc__ 可能不是字符串
    """
    if docstring is None:
        return

    # 如果 docstring 不是字符串，转换为字符串或跳过
    if not isinstance(docstring, str):
        try:
            docstring = str(docstring)
        except:
            # 如果无法转换，使用默认值
            docstring = None

    if docstring:
        try:
            obj.__doc__ = docstring
        except (AttributeError, TypeError):
            # 某些对象不允许设置 __doc__
            pass

def install_hook():
    """在 numpy 加载前安装 hook"""
    # 导入 numpy.core.overrides 前拦截
    import numpy.core.overrides as overrides

    # 替换 add_docstring 函数
    global _original_add_docstring
    _original_add_docstring = overrides.add_docstring
    overrides.add_docstring = safe_add_docstring

# 在 numpy 导入时执行
if 'numpy' not in sys.modules:
    # numpy 还未加载，设置导入钩子
    original_import = builtins.__import__

    def numpy_import_hook(name, *args, **kwargs):
        module = original_import(name, *args, **kwargs)

        # 当 numpy.core.overrides 被导入时，立即修复
        if name == 'numpy.core.overrides' or (
            name.startswith('numpy.core') and hasattr(module, 'add_docstring')
        ):
            if hasattr(module, 'add_docstring') and module.add_docstring != safe_add_docstring:
                global _original_add_docstring
                _original_add_docstring = module.add_docstring
                module.add_docstring = safe_add_docstring
                print("[numpy-fix] Applied safe add_docstring to", name)

        return module

    builtins.__import__ = numpy_import_hook
    print("[numpy-fix] Runtime hook installed")
else:
    # numpy 已经加载，直接修复
    try:
        import numpy.core.overrides
        install_hook()
        print("[numpy-fix] Applied safe add_docstring (numpy already loaded)")
    except:
        pass
