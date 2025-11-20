# -*- coding: utf-8 -*-
"""
更新执行模块
负责执行应用程序的热更新操作
"""

import os
import sys
import time
import subprocess
import shutil
import tempfile
from typing import Optional

from .config import (
    get_app_executable_path,
    get_executable_dir
)
from .config import get_config
from .backup_manager import BackupManager

# 异常类定义
class UpdateExecutionError(Exception):
    """更新执行异常"""
    pass

class UpdateExecutor:
    """更新执行器"""

    def __init__(self):
        self.config = get_config()
        self.backup_manager = BackupManager()

    def execute_update(self, update_file_path: str, new_version: str) -> bool:
        """
        执行应用程序更新
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            # 验证更新文件
            if not os.path.exists(update_file_path):
                raise UpdateExecutionError("更新文件不存在")

            if os.path.getsize(update_file_path) == 0:
                raise UpdateExecutionError("更新文件无效")

            # 获取当前可执行文件路径
            current_exe_path = get_app_executable_path()

            # 如果是开发环境，直接替换文件
            if not getattr(sys, 'frozen', False):
                return self._update_development_environment(update_file_path, new_version)

            # 生产环境（打包后的exe）需要特殊处理
            return self._update_production_environment(update_file_path, new_version)

        except UpdateExecutionError:
            raise
        except Exception as e:
            raise UpdateExecutionError(f"执行更新失败: {str(e)}")

    def _update_development_environment(self, update_file_path: str, new_version: str) -> bool:
        """
        开发环境下的更新（更新版本号并启动新版本）
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            print(f"开发环境更新: 版本 {new_version}")

            # 更新版本文件
            success = self.config.update_current_version(new_version)
            if not success:
                raise UpdateExecutionError("更新版本文件失败")

            print(f"版本已更新到: {new_version}")

            # 在开发环境下，启动新版本程序实例
            try:
                import subprocess
                import sys
                from .config import get_app_executable_path

                current_exe_path = get_app_executable_path()

                # 启动新版本程序
                print("正在启动新版本程序...")
                subprocess.Popen([sys.executable, current_exe_path],
                               cwd=os.path.dirname(current_exe_path),
                               creationflags=subprocess.DETACHED_PROCESS)

                print("新版本程序已启动")

                # 提示用户关闭旧版本
                print("请手动关闭当前程序窗口以完成更新")

                return True

            except Exception as start_error:
                print(f"启动新版本失败: {str(start_error)}")
                # 即使启动失败，版本更新仍然成功
                return True

        except Exception as e:
            raise UpdateExecutionError(f"开发环境更新失败: {str(e)}")

    def _update_production_environment(self, update_file_path: str, new_version: str) -> bool:
        """
        生产环境下的更新（需要处理文件占用）
        :param update_file_path: 更新文件路径
        :param new_version: 新版本号
        :return: 是否更新成功
        """
        try:
            current_exe_path = get_app_executable_path()

            # 创建备份
            backup_path = self.backup_manager.create_backup()
            if not backup_path:
                raise UpdateExecutionError("创建备份失败")

            print(f"已创建备份: {backup_path}")

            # 尝试替换可执行文件
            if not self._replace_executable(update_file_path, current_exe_path):
                # 如果直接替换失败，使用批处理脚本延迟更新
                return self._schedule_delayed_update(update_file_path, current_exe_path, new_version)

            # 更新版本文件
            self.config.update_current_version(new_version)

            print("文件替换成功，更新完成")
            return True

        except Exception as e:
            raise UpdateExecutionError(f"生产环境更新失败: {str(e)}")

    def _replace_executable(self, source_path: str, target_path: str) -> bool:
        """
        替换可执行文件
        :param source_path: 源文件路径
        :param target_path: 目标文件路径
        :return: 是否替换成功
        """
        try:
            # 等待文件释放
            for _ in range(10):  # 最多等待10秒
                try:
                    # 尝试删除目标文件
                    if os.path.exists(target_path):
                        os.remove(target_path)

                    # 复制新文件
                    shutil.copy2(source_path, target_path)

                    # 验证文件是否正确复制
                    if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                        return True

                except PermissionError:
                    print("文件被占用，等待释放...")
                    time.sleep(1)
                except Exception as e:
                    print(f"替换文件失败: {e}")
                    time.sleep(1)

            return False

        except Exception as e:
            print(f"替换可执行文件失败: {e}")
            return False

    def _schedule_delayed_update(self, update_file_path: str, current_exe_path: str, new_version: str) -> bool:
        """
        安排延迟更新（使用增强的批处理脚本）
        :param update_file_path: 更新文件路径
        :param current_exe_path: 当前可执行文件路径
        :param new_version: 新版本号
        :return: 是否成功安排延迟更新
        """
        try:
            # 预检查：验证文件和路径
            if not os.path.exists(update_file_path):
                raise UpdateExecutionError(f"更新文件不存在: {update_file_path}")

            if os.path.getsize(update_file_path) == 0:
                raise UpdateExecutionError(f"更新文件为空: {update_file_path}")

            # 检查目标目录是否可写
            target_dir = os.path.dirname(current_exe_path)
            if not os.access(target_dir, os.W_OK):
                raise UpdateExecutionError(f"目标目录无写入权限: {target_dir}")

            # 创建日志文件路径
            log_path = os.path.join(tempfile.gettempdir(), "update_log.txt")

            # 创建增强的更新脚本
            script_content = f'''@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========== 开始更新应用程序 ==========
echo 更新时间: %date% %time% >> "{log_path}"
echo 源文件: "{update_file_path}" >> "{log_path}"
echo 目标文件: "{current_exe_path}" >> "{log_path}"

REM 预检查：验证源文件存在
if not exist "{update_file_path}" (
    echo 错误：源文件不存在！ >> "{log_path}"
    echo 错误：源文件不存在！
    timeout /t 5 /nobreak >nul
    exit /b 1
)

REM 预检查：验证源文件大小
for %%A in ("{update_file_path}") do set size=%%~zA
if !size! LEQ 0 (
    echo 错误：源文件为空！ >> "{log_path}"
    echo 错误：源文件为空！
    timeout /t 5 /nobreak >nul
    exit /b 1
)

echo 源文件检查通过，大小：!size! 字节 >> "{log_path}"

REM 等待原程序退出（最多等待30秒）
echo 等待原程序退出...
set wait_count=0
:wait_exit
timeout /t 1 /nobreak >nul
set /a wait_count+=1

REM 检查目标文件是否可访问（即原程序是否已退出）
if exist "{current_exe_path}" (
    REM 尝试重命名文件来测试是否被占用
    ren "{current_exe_path}" "Sap_Operate_theme_backup.exe" >nul 2>&1
    if !errorlevel! EQU 0 (
        ren "Sap_Operate_theme_backup.exe" "{current_exe_path}" >nul 2>&1
        echo 原程序已退出，文件可访问 >> "{log_path}"
        goto file_replace
    )

    if !wait_count! GEQ 30 (
        echo 警告：等待超时，强制继续更新 >> "{log_path}"
        echo 警告：等待超时，强制继续更新
        goto file_replace
    )

    echo 等待中...(!wait_count!/30秒) >> "{log_path}"
    goto wait_exit
) else (
    echo 目标文件不存在，继续更新 >> "{log_path}"
    goto file_replace
)

:file_replace
echo 开始文件替换操作...

REM 创建备份
if exist "{current_exe_path}" (
    echo 创建备份文件 >> "{log_path}"
    copy "{current_exe_path}" "{current_exe_path}.backup" >nul 2>&1
)

REM 文件替换操作（带重试机制）
set retry_count=0
:max_retry
echo 尝试复制文件 (第!retry_count!次) >> "{log_path}"

REM 删除目标文件（如果存在）
if exist "{current_exe_path}" (
    del "{current_exe_path}" >nul 2>&1
    if !errorlevel! NEQ 0 (
        echo 删除目标文件失败，等待重试... >> "{log_path}"
        timeout /t 2 /nobreak >nul
        set /a retry_count+=1
        if !retry_count! LSS 3 goto max_retry
        echo 错误：删除目标文件失败，重试次数已用尽！ >> "{log_path}"
        echo 错误：删除目标文件失败！
        timeout /t 5 /nobreak >nul
        exit /b 1
    )
)

REM 复制新文件
copy /Y "{update_file_path}" "{current_exe_path}" >nul 2>&1
if !errorlevel! NEQ 0 (
    echo 复制文件失败，等待重试... >> "{log_path}"
    timeout /t 2 /nobreak >nul
    set /a retry_count+=1
    if !retry_count! LSS 3 goto max_retry
    echo 错误：复制文件失败，重试次数已用尽！ >> "{log_path}"
    echo 错误：复制文件失败！
    timeout /t 5 /nobreak >nul
    exit /b 1
)

echo 文件复制成功 >> "{log_path}"

REM 验证复制结果
if exist "{current_exe_path}" (
    for %%A in ("{current_exe_path}") do set new_size=%%~zA
    if !new_size! EQU !size! (
        echo 文件验证成功，大小匹配：!new_size! 字节 >> "{log_path}"
        echo 文件替换成功！
    ) else (
        echo 错误：文件大小不匹配！期望：!size!，实际：!new_size! >> "{log_path}"
        echo 错误：文件大小不匹配！
        timeout /t 5 /nobreak >nul
        exit /b 1
    )
) else (
    echo 错误：目标文件不存在！ >> "{log_path}"
    echo 错误：目标文件不存在！
    timeout /t 5 /nobreak >nul
    exit /b 1
)

REM 等待1秒确保文件完全写入
timeout /t 1 /nobreak >nul

REM 启动新版本
echo 启动新版本应用程序... >> "{log_path}"
start "" "{current_exe_path}"

REM 等待启动验证
timeout /t 3 /nobreak >nul

REM 检查新版本是否启动成功
tasklist /FI "IMAGENAME eq Sap_Operate_theme.exe" 2>NUL | find /I "Sap_Operate_theme.exe" >NUL
if !errorlevel! EQU 0 (
    echo 新版本启动成功 >> "{log_path}"
    echo 新版本启动成功！
) else (
    echo 警告：无法确认新版本启动状态 >> "{log_path}"
    echo 警告：无法确认新版本启动状态
)

REM 清理临时文件
echo 清理临时文件... >> "{log_path}"
del "{update_file_path}" >nul 2>&1
del "%~f0" >nul 2>&1

echo ========== 更新完成 ========== >> "{log_path}"
echo 更新完成！新版本已启动。
echo 详细日志：{log_path}
timeout /t 5 /nobreak >nul
exit /b 0
'''

            # 创建临时脚本文件
            script_path = os.path.join(tempfile.gettempdir(), "enhanced_update_script.bat")
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)

            # 清理旧的日志文件
            if os.path.exists(log_path):
                try:
                    os.remove(log_path)
                except Exception:
                    pass

            print(f"已创建增强更新脚本: {script_path}")
            print(f"更新日志将保存到: {log_path}")

            # 启动更新脚本
            subprocess.Popen([script_path],
                           creationflags=subprocess.DETACHED_PROCESS,
                           env=os.environ.copy(),
                           encoding='utf-8')

            print("已安排延迟更新，应用程序将重启")
            print("更新过程包含详细日志记录，如遇问题可查看日志文件")
            return True

        except Exception as e:
            raise UpdateExecutionError(f"安排延迟更新失败: {str(e)}")

    def restart_application(self) -> bool:
        """
        重启应用程序
        :return: 是否成功重启
        """
        try:
            if getattr(sys, 'frozen', False):
                # 打包后的exe
                current_exe = sys.executable
                subprocess.Popen([current_exe],
                               env=os.environ.copy(),
                               encoding='utf-8')
            else:
                # 开发环境
                current_script = sys.argv[0]
                subprocess.Popen([sys.executable, current_script],
                               env=os.environ.copy(),
                               encoding='utf-8')

            # 退出当前进程
            sys.exit(0)

        except Exception as e:
            print(f"重启应用程序失败: {e}")
            return False

    def rollback_update(self) -> bool:
        """
        回滚到上一个版本
        :return: 是否回滚成功
        """
        try:
            # 获取最新备份
            latest_backup = self.backup_manager.get_latest_backup()
            if not latest_backup:
                raise UpdateExecutionError("没有找到可用的备份文件")

            # 从备份恢复
            success = self.backup_manager.restore_from_backup(latest_backup)
            if not success:
                raise UpdateExecutionError("从备份恢复失败")

            # 更新版本文件
            # 注意：这里需要根据实际情况确定如何获取备份的版本号
            # 暂时使用本地版本管理器的当前版本

            print("回滚成功")
            return True

        except Exception as e:
            raise UpdateExecutionError(f"回滚失败: {str(e)}")

    def validate_update_file(self, update_file_path: str) -> tuple:
        """
        验证更新文件的有效性
        :param update_file_path: 更新文件路径
        :return: (是否有效, 错误信息)
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(update_file_path):
                return False, "更新文件不存在"

            # 检查文件大小
            file_size = os.path.getsize(update_file_path)
            if file_size == 0:
                return False, "更新文件为空"

            # 检查文件扩展名
            if not (update_file_path.endswith('.exe') or update_file_path.endswith('.zip')):
                return False, "更新文件格式不正确"

            # 基本可执行文件检查
            if update_file_path.endswith('.exe'):
                try:
                    # 检查PE文件头（简单检查）
                    with open(update_file_path, 'rb') as f:
                        header = f.read(2)
                        if header != b'MZ':  # DOS header
                            return False, "更新文件不是有效的可执行文件"
                except Exception as e:
                    return False, f"读取更新文件失败: {str(e)}"

            return True, "更新文件有效"

        except Exception as e:
            return False, f"验证更新文件失败: {str(e)}"

    def get_update_progress_info(self) -> dict:
        """
        获取更新进度信息
        :return: 进度信息字典
        """
        return {
            'is_updating': False,  # 是否正在更新
            'current_step': '',    # 当前步骤
            'progress': 0,         # 进度百分比
            'error_message': ''    # 错误信息
        }