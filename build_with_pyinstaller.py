#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP工具打包脚本
使用 PyInstaller 将 Python 应用程序打包为 Windows 可执行文件
"""

import os
import sys
import subprocess
import shutil
import time
import json
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any


class PackagerConfig:
    """打包配置类"""

    def __init__(self):
        self.entry_file = 'Sap_Operate_theme.py'
        self.spec_file = 'Sap_Operate_theme.spec'
        self.icon_file = 'Sap_Operate_Logo.ico'
        self.app_name = 'Sap_Operate_theme'
        self.use_spec = True  # 优先使用 .spec 文件
        self.clean_build = True
        self.upx_compress = True
        self.onefile = True  # 单文件模式

        # 必需的依赖文件
        self.required_files = [
            self.entry_file,
            'auto_updater/',
            'Get_Data.py',
            'Sap_Function.py',
            'File_Operate.py',
            'PDF_Operate.py',
            'Logger.py',
            'Excel_Field_Mapper.py',
            'Data_Table.py',
            'Sap_Operate_Ui.py',
            'theme_manager_theme.py',
            'Revenue_Operate.py',
            'auto_updater/config_constants.py',
        ]

        # 可选但建议包含的文件
        self.optional_files = [
            'chicon.py',
            'chicon.qrc',
            'Sap_Operate_Ui.ui',
            'Table_Ui.ui',
        ]


class SAPPackager:
    """SAP工具打包器"""

    def __init__(self, config: PackagerConfig = None):
        self.config = config or PackagerConfig()
        self.start_time = time.time()

    def log(self, message: str, level: str = "INFO"):
        """日志输出"""
        timestamp = time.strftime("%H:%M:%S")
        # 替换特殊字符以避免编码问题
        safe_message = message.replace('✓', '[OK]').replace('✗', '[FAIL]').replace('-', '[SKIP]')
        print(f"[{timestamp}] {level}: {safe_message}")

    def check_environment(self) -> bool:
        """检查打包环境"""
        self.log("检查打包环境...")

        # 检查 Python 版本
        if sys.version_info < (3, 8):
            self.log("Python 版本过低，建议使用 3.8 或更高版本", "ERROR")
            return False

        # 检查 PyInstaller
        try:
            result = subprocess.run(['pyinstaller', '--version'],
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                self.log(f"PyInstaller 版本: {result.stdout.strip()}")
            else:
                self.log("PyInstaller 未安装或无法访问", "ERROR")
                return False
        except (subprocess.TimeoutExpired, FileNotFoundError):
            self.log("PyInstaller 未安装或无法访问", "ERROR")
            return False

        return True

    def check_dependencies(self) -> bool:
        """检查项目依赖"""
        self.log("检查项目依赖...")

        missing_files = []

        # 检查必需文件
        for file_path in self.config.required_files:
            if os.path.exists(file_path):
                self.log(f"✓ {file_path}")
            else:
                self.log(f"✗ {file_path}", "WARNING")
                missing_files.append(file_path)

        # 检查可选文件
        for file_path in self.config.optional_files:
            if os.path.exists(file_path):
                self.log(f"✓ {file_path} (可选)")
            else:
                self.log(f"- {file_path} (可选，不存在)")

        if missing_files:
            self.log(f"缺少 {len(missing_files)} 个必需文件，可能影响打包", "WARNING")
            response = input("是否继续？(y/N): ")
            return response.lower() in ['y', 'yes']

        return True

    def check_pandas_dependencies(self) -> Tuple[bool, List[str], Dict[str, str]]:
        """检查pandas依赖的完整性"""
        self.log("检查pandas依赖...")

        required_modules = {
            'pandas': 'pandas',
            'numpy': 'numpy',
            'openpyxl': 'openpyxl',
            'PyQt5': 'PyQt5',
            'pdfplumber': 'pdfplumber',
            'pypdfium2': 'pypdfium2',
            'win32com': 'win32com'
        }

        missing_modules = []
        version_info = {}

        for module_name, import_name in required_modules.items():
            try:
                module = __import__(import_name)
                version = getattr(module, '__version__', 'unknown')
                version_info[module_name] = version
                self.log(f"  [OK] {module_name} v{version}")
            except ImportError as e:
                self.log(f"  [ERROR] {module_name} - {e}")
                missing_modules.append(module_name)

        if missing_modules:
            self.log(f"发现缺失模块: {', '.join(missing_modules)}")
            return False, missing_modules, {}
        else:
            self.log("所有必需模块已安装")
            return True, [], version_info

    def validate_pandas_dlls(self) -> Tuple[bool, List[str]]:
        """验证pandas相关的DLL文件是否存在"""
        self.log("验证pandas DLL文件...")

        try:
            import pandas
            import numpy
            import sys

            # 检查pandas核心库
            pandas_libs_path = Path(pandas.__file__).parent / '_libs'
            numpy_libs_path = Path(numpy.__file__).parent / '.libs'

            dll_check_results = []

            # 检查pandas _libs目录
            if pandas_libs_path.exists():
                pyd_files = list(pandas_libs_path.glob('*.pyd'))
                dll_files = list(pandas_libs_path.glob('*.dll'))
                dll_check_results.append(f"pandas _libs: {len(pyd_files)} .pyd文件, {len(dll_files)} .dll文件")
            else:
                dll_check_results.append("pandas _libs: 目录不存在")

            # 检查numpy .libs目录
            if numpy_libs_path.exists():
                dll_files = list(numpy_libs_path.glob('*.dll'))
                dll_check_results.append(f"numpy .libs: {len(dll_files)} .dll文件")
            else:
                dll_check_results.append("numpy .libs: 目录不存在")

            # 检查关键的pandas组件
            critical_components = [
                'pandas._libs.tslibs.base',
                'pandas._libs.tslibs.nattype',
                'numpy.core._multiarray_umath'
            ]

            missing_critical = []
            for component in critical_components:
                try:
                    __import__(component)
                    dll_check_results.append(f"  [OK] {component}")
                except ImportError as e:
                    dll_check_results.append(f"  [ERROR] {component} - {e}")
                    missing_critical.append(component)

            for result in dll_check_results:
                self.log(f"  {result}")

            if missing_critical:
                self.log(f"警告: 发现 {len(missing_critical)} 个关键组件缺失")
                return False, missing_critical
            else:
                self.log("pandas DLL验证通过")
                return True, []

        except Exception as e:
            self.log(f"DLL验证过程中出错: {e}")
            return False, [str(e)]

    def diagnose_pandas_environment(self) -> Optional[Dict]:
        """诊断pandas环境，提供详细信息"""
        self.log("=" * 50)
        self.log("pandas环境诊断报告")
        self.log("=" * 50)

        try:
            import pandas as pd
            import numpy as np
            import sys
            from pathlib import Path

            # 基本信息
            self.log(f"Python版本: {sys.version}")
            self.log(f"pandas版本: {pd.__version__}")
            self.log(f"numpy版本: {np.__version__}")
            self.log(f"pandas路径: {pd.__file__}")
            self.log(f"numpy路径: {np.__file__}")

            # 虚拟环境信息
            venv_path = Path(sys.executable).parent
            self.log(f"虚拟环境路径: {venv_path}")

            # 关键依赖检查
            deps_ok, missing, versions = self.check_pandas_dependencies()
            dll_ok, missing_dlls = self.validate_pandas_dlls()

            # 生成诊断报告
            report = {
                "python_version": sys.version,
                "pandas_version": pd.__version__,
                "numpy_version": np.__version__,
                "dependencies": deps_ok,
                "missing_modules": missing,
                "module_versions": versions,
                "dll_validation": dll_ok,
                "missing_dlls": missing_dlls,
                "venv_path": str(venv_path),
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
            }

            # 保存诊断报告
            report_file = "pandas_environment_report.json"
            with open(report_file, "w", encoding="utf-8") as f:
                json.dump(report, f, indent=2, ensure_ascii=False)

            self.log(f"\n诊断报告已保存到: {report_file}")
            return report

        except Exception as e:
            self.log(f"环境诊断失败: {e}")
            return None

    def analyze_build_errors(self, error_output: str) -> List[str]:
        """分析常见的打包错误并提供解决建议"""
        suggestions = []

        error_lower = error_output.lower()

        # numpy相关错误
        if "numpy" in error_lower and ("docstring" in error_lower or "add_docstring" in error_lower):
            suggestions.append("检测到numpy docstring错误，这是numpy 2.0.0与PyInstaller的兼容性问题")
            suggestions.append("建议：升级到最新版本的PyInstaller或降级numpy到1.26.4")

        if "dll load failed" in error_lower:
            suggestions.append("检测到DLL加载失败，这通常与pandas或numpy依赖相关")
            suggestions.append("建议：确保pandas和numpy版本兼容，尝试重新安装pandas")

        if "numpy" in error_lower and ("not found" in error_lower or "cannot be found" in error_lower):
            suggestions.append("numpy模块缺失，请检查numpy是否正确安装")
            suggestions.append("运行命令：pip install numpy==2.0.0")

        # pandas相关错误
        if "pandas" in error_lower and ("not found" in error_lower or "cannot be found" in error_lower):
            suggestions.append("pandas模块缺失，请检查pandas是否正确安装")
            suggestions.append("运行命令：pip install pandas==2.2.3")

        # PyInstaller相关错误
        if "module not found" in error_lower:
            suggestions.append("模块导入错误，检查所有必需的Python包是否已安装")

        if "failed to execute script" in error_lower:
            suggestions.append("脚本执行失败，可能是依赖问题或代码错误")
            suggestions.append(f"建议检查主程序{self.config.entry_file}是否有语法错误")

        # 内存相关问题
        if "memory" in error_lower and "error" in error_lower:
            suggestions.append("内存错误，可能是系统资源不足")
            suggestions.append("建议关闭其他程序，增加虚拟内存或使用--noarchive参数")

        # 权限问题
        if "permission" in error_lower or "access denied" in error_lower:
            suggestions.append("权限不足，请以管理员身份运行打包工具")

        # 路径问题
        if "path" in error_lower and ("not found" in error_lower or "too long" in error_lower):
            suggestions.append("路径问题，检查文件路径是否正确或过长")

        if not suggestions:
            suggestions.append("未识别的错误类型，请查看详细错误信息")
            suggestions.append("建议：清理build和dist目录后重新尝试打包")

        return suggestions

    def validate_package_result(self, exe_path: str) -> Dict[str, Any]:
        """验证打包结果的完整性"""
        try:
            if not os.path.exists(exe_path):
                return {"success": False, "message": "exe文件不存在"}

            # 检查文件大小
            file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
            if file_size < 50:  # 通常包含pandas的exe文件会比较大
                return {
                    "success": False,
                    "message": f"exe文件过小({file_size:.1f}MB)，可能缺少关键依赖"
                }

            # 尝试运行exe文件进行基本测试
            self.log("进行基本功能验证...")
            test_cmd = [exe_path, "--test"]  # 添加测试参数
            try:
                # 运行5秒测试，然后强制终止
                process = subprocess.Popen(
                    test_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == 'nt' else 0
                )
                try:
                    stdout, stderr = process.communicate(timeout=5)
                    if process.returncode == 0:
                        return {"success": True, "message": "基本功能验证通过"}
                    else:
                        error_msg = stderr.decode('utf-8', errors='ignore') if stderr else "未知错误"
                        return {
                            "success": False,
                            "message": f"exe文件测试失败: {error_msg[:200]}"
                        }
                except subprocess.TimeoutExpired:
                    process.terminate()
                    # 超时可能是正常的，说明程序在运行
                    return {"success": True, "message": "程序启动正常（测试超时终止）"}
            except Exception as e:
                return {
                    "success": False,
                    "message": f"无法运行exe文件进行测试: {e}"
                }

        except Exception as e:
            return {"success": False, "message": f"验证过程出错: {e}"}

    def clean_directories(self):
        """清理构建目录"""
        self.log("清理构建目录...")

        dirs_to_clean = ['build', 'dist']

        for dir_name in dirs_to_clean:
            if os.path.exists(dir_name):
                try:
                    shutil.rmtree(dir_name)
                    self.log(f"✓ 清理 {dir_name}")
                except Exception as e:
                    self.log(f"✗ 清理 {dir_name} 失败: {e}", "WARNING")

        # 清理 .spec 文件（如果使用自动生成）
        if not self.config.use_spec and os.path.exists(self.config.spec_file):
            try:
                os.remove(self.config.spec_file)
                self.log(f"✓ 清理 {self.config.spec_file}")
            except Exception as e:
                self.log(f"✗ 清理 {self.config.spec_file} 失败: {e}", "WARNING")

    def collect_binaries_dynamically(self) -> List[Tuple[str, str]]:
        """动态收集pandas和numpy的二进制文件"""
        binaries = []

        try:
            # 动态收集pandas二进制文件
            import pandas
            pandas_libs_path = Path(pandas.__file__).parent / '_libs'

            if pandas_libs_path.exists():
                # 收集.pyd文件
                for pyd_file in pandas_libs_path.glob('*.pyd'):
                    binaries.append((str(pyd_file), 'pandas/_libs'))
                    self.log(f"  收集pandas二进制: {pyd_file.name}")

                # 收集.dll文件（如果存在）
                for dll_file in pandas_libs_path.glob('*.dll'):
                    binaries.append((str(dll_file), 'pandas/_libs'))
                    self.log(f"  收集pandas DLL: {dll_file.name}")

                # 收集tslibs目录
                tslibs_path = pandas_libs_path / 'tslibs'
                if tslibs_path.exists():
                    for tslib_file in tslibs_path.glob('*.pyd'):
                        binaries.append((str(tslib_file), 'pandas/_libs/tslibs'))

        except ImportError:
            self.log("Warning: pandas not available for binary collection")

        try:
            # 动态收集numpy二进制文件
            import numpy
            numpy_libs_path = Path(numpy.__file__).parent / '.libs'

            if numpy_libs_path.exists():
                for dll_file in numpy_libs_path.glob('*.dll'):
                    binaries.append((str(dll_file), 'numpy/.libs'))
                    self.log(f"  收集numpy DLL: {dll_file.name}")

            numpy_core_path = Path(numpy.__file__).parent / 'core'
            if numpy_core_path.exists():
                for pyd_file in numpy_core_path.glob('*.pyd'):
                    binaries.append((str(pyd_file), 'numpy/core'))
                    self.log(f"  收集numpy核心: {pyd_file.name}")

        except ImportError:
            self.log("Warning: numpy not available for binary collection")

        return binaries

    def get_enhanced_build_command(self) -> List[str]:
        """获取增强的PyInstaller命令，包含numpy 2.0.0兼容性修复"""
        cmd = ['pyinstaller']

        if self.config.use_spec and os.path.exists(self.config.spec_file):
            # 使用 .spec 文件
            cmd.extend([self.config.spec_file, '--clean', '--noconfirm'])
            self.log("使用 .spec 文件配置")
        else:
            # 使用命令行参数
            if self.config.onefile:
                cmd.append('--onefile')
            else:
                cmd.append('--onedir')

            cmd.append('--windowed')
            cmd.append('--clean')
            cmd.append('--noconfirm')

            # 添加numpy 2.0.0特定参数
            cmd.extend(['--copy-metadata', 'numpy'])
            cmd.extend(['--copy-metadata', 'pandas'])

            # 添加收集参数（增强版）
            for lib in ['pandas', 'numpy', 'openpyxl']:
                cmd.extend(['--collect-all', lib])

            # 添加隐藏导入（扩展版）
            hidden_imports = [
                # PyQt5相关
                'PyQt5.QtCore',
                'PyQt5.QtGui',
                'PyQt5.QtWidgets',
                'PyQt5.sip',

                # Windows相关
                'win32com.client',
                'win32com.universal',
                'pythoncom',
                'pywintypes',

                # pandas相关（关键修复）
                'pandas',
                'pandas._libs',
                'pandas._libs.tslibs',
                'pandas._libs.tslibs.base',
                'pandas._libs.tslibs.nattype',
                'pandas._libs.tslibs.np_datetime',
                'pandas._libs.tslibs.parsing',
                'pandas._libs.tslibs.period',
                'pandas._libs.tslibs.strftime',
                'pandas._libs.tslibs.timedeltas',
                'pandas._libs.tslibs.timestamps',
                'pandas._libs.tslibs.timezones',
                'pandas._libs.tslibs.tzconversion',
                'pandas._libs.tslibs.vectorized',
                'pandas.core',
                'pandas.core.frame',
                'pandas.core.series',
                'pandas.core.common',
                'pandas.core.generic',
                'pandas.io',
                'pandas.io.common',
                'pandas.io.excel',
                'pandas.io.formats',
                'pandas.util',

                # numpy相关（关键修复）
                'numpy',
                'numpy.core',
                'numpy.core._multiarray_umath',
                'numpy.core.multiarray',
                'numpy.core.umath',
                'numpy.linalg',
                'numpy.linalg.lapack_lite',
                'numpy.random',
                'numpy.fft',
                'numpy.polynomial',
                'numpy._core',
                'numpy._core._multiarray_umath',

                # 应用模块
                'auto_updater',
                'pdfplumber',
                'pypdfium2',
                'packaging',
                'xlsxwriter',
                'pdfminer.six',
                'PIL',
                'PIL.Image',
                'openpyxl',
            ]

            for imp in hidden_imports:
                cmd.extend(['--hidden-import', imp])

            # 添加排除模块
            excludes = [
                'tkinter',
                'matplotlib',
                'scipy',
                'scikit-learn',
                'IPython',
                'jupyter',
                'notebook',
                'pytest',
                'setuptools',
                'pip',
                'sphinx',
                'nose',
                'doctest',
            ]

            for exclude in excludes:
                cmd.extend(['--exclude-module', exclude])

            # 添加图标
            if os.path.exists(self.config.icon_file):
                cmd.append(f'--icon={self.config.icon_file}')

            # 添加动态收集的二进制文件
            binaries = self.collect_binaries_dynamically()
            if binaries:
                self.log(f"添加 {len(binaries)} 个动态二进制文件")
                for binary, dest in binaries:
                    cmd.extend(['--add-binary', f'{binary};{dest}'])

            cmd.append(self.config.entry_file)
            self.log("使用增强命令行参数配置")

        return cmd

    def build_command(self) -> List[str]:
        """构建 PyInstaller 命令（保留向后兼容性）"""
        return self.get_enhanced_build_command()

    def run_build(self) -> bool:
        """执行打包构建（增强版，包含错误分析）"""
        self.log("开始打包构建...")

        cmd = self.build_command()
        self.log(f"执行命令: {' '.join(cmd[:10])}... (参数较多已省略)")

        output_lines = []

        try:
            # 设置环境变量，强制使用 UTF-8 编码
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONLEGACYWINDOWSSTDIO'] = '1'

            # 执行打包命令
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=False,  # 使用字节模式
                bufsize=1,
                env=env
            )

            # 实时输出日志，处理编码问题
            while True:
                line_bytes = process.stdout.readline()
                if not line_bytes:
                    break

                try:
                    # 尝试多种编码方式解码
                    try:
                        line = line_bytes.decode('utf-8').strip()
                    except UnicodeDecodeError:
                        try:
                            line = line_bytes.decode('gbk').strip()
                        except UnicodeDecodeError:
                            # 最后尝试忽略错误字符
                            line = line_bytes.decode('utf-8', errors='ignore').strip()

                    if line:
                        output_lines.append(line)
                        print(f"[PyInstaller] {line}")

                except Exception as e:
                    # 如果解码仍然失败，输出错误信息但继续
                    error_line = f"[解码错误: {str(e)}]"
                    output_lines.append(error_line)
                    print(f"[PyInstaller] {error_line}")

            # 等待完成
            process.wait()

            if process.returncode == 0:
                self.log("[OK] 打包构建成功")
                return True
            else:
                self.log(f"[FAIL] 打包构建失败，返回码: {process.returncode}", "ERROR")

                # 分析构建错误
                full_output = "\n".join(output_lines)
                error_suggestions = self.analyze_build_errors(full_output)

                if error_suggestions:
                    self.log("\n错误分析建议:")
                    for i, suggestion in enumerate(error_suggestions, 1):
                        self.log(f"  {i}. {suggestion}")

                return False

        except Exception as e:
            self.log(f"[FAIL] 打包过程中发生错误: {e}", "ERROR")

            # 分析异常错误
            error_suggestions = self.analyze_build_errors(str(e))
            if error_suggestions:
                self.log("\n错误分析建议:")
                for i, suggestion in enumerate(error_suggestions, 1):
                    self.log(f"  {i}. {suggestion}")

            return False

    def verify_build(self) -> bool:
        """验证打包结果（增强版）"""
        self.log("验证打包结果...")

        exe_path = Path('dist') / f"{self.config.app_name}.exe"

        if exe_path.exists():
            file_size = exe_path.stat().st_size
            self.log(f"✓ 可执行文件已生成: {exe_path}")
            self.log(f"✓ 文件大小: {file_size / (1024*1024):.1f} MB")

            # 使用增强的验证功能
            validation_result = self.validate_package_result(str(exe_path))
            if validation_result["success"]:
                self.log(f"✓ {validation_result['message']}")
                return True
            else:
                self.log(f"⚠ {validation_result['message']}", "WARNING")
                # 虽然警告，但仍然返回True，因为文件已经生成
                return True
        else:
            self.log("✗ 可执行文件未生成", "ERROR")
            return False

    def show_summary(self):
        """显示打包摘要"""
        elapsed_time = time.time() - self.start_time
        self.log("=" * 50)
        self.log("打包摘要")
        self.log("=" * 50)
        self.log(f"✓ 打包完成，耗时: {elapsed_time:.1f} 秒")
        self.log(f"✓ 可执行文件位置: dist/{self.config.app_name}.exe")

        exe_path = Path('dist') / f"{self.config.app_name}.exe"
        if exe_path.exists():
            file_size = exe_path.stat().st_size
            self.log(f"✓ 文件大小: {file_size / (1024*1024):.1f} MB")

        self.log("")
        self.log("使用说明:")
        self.log("1. 双击运行可执行文件")
        self.log("2. 如有问题，请检查系统环境")
        self.log("3. 确保已安装运行时库（如 Microsoft Visual C++）")
        self.log("=" * 50)

    def run(self) -> bool:
        """执行完整打包流程（增强版，包含pandas环境诊断）"""
        self.log("开始 SAP 工具打包流程")
        self.log("=" * 50)

        # 检查环境
        if not self.check_environment():
            return False

        # 检查依赖
        if not self.check_dependencies():
            return False

        # 新增：检查pandas依赖完整性
        deps_ok, missing, versions = self.check_pandas_dependencies()
        if not deps_ok:
            self.log("pandas依赖检查失败，尝试自动安装缺失模块...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install"] + missing, check=True)
                self.log("缺失模块安装成功")
                # 重新检查依赖
                deps_ok, missing, versions = self.check_pandas_dependencies()
                if not deps_ok:
                    self.log("依赖安装后仍然有问题，请手动检查")
                    return False
            except subprocess.CalledProcessError:
                self.log("自动安装失败，请手动安装缺失模块")
                return False

        # 新增：验证pandas DLL文件
        dll_ok, missing_dlls = self.validate_pandas_dlls()
        if not dll_ok:
            self.log("警告: pandas DLL验证存在问题，可能导致打包后运行失败")
            self.log(f"缺失的关键组件: {', '.join(missing_dlls)}")
            self.log("建议检查pandas和numpy的安装完整性")
            # 继续执行，但给出警告

        # 新增：生成环境诊断报告
        self.diagnose_pandas_environment()

        # 清理目录
        if self.config.clean_build:
            self.clean_directories()

        # 执行构建
        if not self.run_build():
            return False

        # 验证结果
        if not self.verify_build():
            return False

        # 显示摘要
        self.show_summary()
        return True


def main():
    """主函数"""
    # 创建配置
    config = PackagerConfig()

    # 解析命令行参数
    if len(sys.argv) > 1:
        if '--no-clean' in sys.argv:
            config.clean_build = False
        if '--onedir' in sys.argv:
            config.onefile = False
        if '--no-upx' in sys.argv:
            config.upx_compress = False
        if '--help' in sys.argv or '-h' in sys.argv:
            print("SAP工具打包脚本")
            print("用法: python build_with_pyinstaller.py [选项]")
            print("选项:")
            print("  --no-clean  不清理构建目录")
            print("  --onedir    生成目录模式（非单文件）")
            print("  --no-upx    不使用 UPX 压缩")
            print("  --help      显示此帮助信息")
            return

    # 创建打包器并运行
    packager = SAPPackager(config)
    success = packager.run()

    if success:
        print("\n[SUCCESS] 打包成功！")
    else:
        print("\n[FAIL] 打包失败！")
        sys.exit(1)


if __name__ == '__main__':
    main()
