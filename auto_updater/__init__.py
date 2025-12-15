# -*- coding: utf-8 -*-
"""
PDF重命名工具自动更新模块
提供基于GitHub Releases的自动更新功能
"""

from typing import Optional, Callable, Tuple
import logging
from .config import get_config
from .github_client import GitHubClient
from .download_manager import DownloadManager
from .backup_manager import BackupManager
from .update_executor import UpdateExecutor
from .config import *
from .config_constants import (
    ERROR_DOWNLOAD_URL_FAILED,
    ERROR_DOWNLOAD_URL_TITLE,
    NETWORK_TIMEOUT_MEDIUM
)
from enum import Enum

logger = logging.getLogger(__name__)

# 下载状态枚举
class DownloadState(Enum):
    IDLE = "idle"                    # 空闲状态
    DOWNLOADING = "downloading"        # 正在下载
    COMPLETED = "completed"            # 下载完成
    FAILED = "failed"                  # 下载失败
    CANCELLED = "cancelled"            # 用户取消

# 自定义异常类
class UpdateError(Exception):
    """更新功能基础异常"""
    pass

class NetworkError(UpdateError):
    """网络连接异常"""
    pass

class VersionCheckError(UpdateError):
    """版本检查异常"""
    pass

class DownloadError(UpdateError):
    """文件下载异常"""
    pass

class BackupError(UpdateError):
    """备份操作异常"""
    pass

class UpdateExecutionError(UpdateError):
    """更新执行异常"""
    pass

# 主要接口类
class AutoUpdater:
    """
    自动更新器主类
    整合所有更新功能组件，提供统一的更新接口
    """

    def __init__(self, parent=None):
        """
        初始化自动更新器
        :param parent: 父对象（用于GUI信号连接）
        """
        self.config = get_config()
        self.github_client = GitHubClient()
        self.download_manager = DownloadManager()
        self.backup_manager = BackupManager()
        self.update_executor = UpdateExecutor()
        self.parent = parent

        # 初始化UI管理器
        try:
            from .ui.update_ui_manager import UpdateUIManager
            self.ui_manager = UpdateUIManager(parent)
        except ImportError:
            self.ui_manager = None

        # 下载状态管理
        self._download_state = DownloadState.IDLE
        self._current_download_version = None  # 当前正在下载的版本
        self._download_confirm_calls = 0        # 防止重复调用计数器

    def check_for_updates(self, force_check=False) -> tuple:
        """
        检查更新
        :param force_check: 是否强制检查（忽略时间间隔）
        :return: (是否有更新, 远程版本, 本地版本, 错误信息)
        """
        try:
            local_version = self.config.current_version

            # 检查是否应该进行更新检查
            if not force_check and not self.config.should_check_for_updates():
                return False, None, local_version, "距离上次检查时间过短"

            # 获取远程版本信息
            release_info = self.github_client.get_latest_release()
            if not release_info:
                return False, None, local_version, "无法获取远程版本信息"

            remote_version = release_info.get('tag_name', '').lstrip('v')
            if not remote_version:
                return False, None, local_version, "远程版本格式无效"

            # 检查是否有更新
            has_update = self.config.is_newer_version(remote_version, local_version)

            # 更新最后检查时间
            self.config.update_last_check_time()

            return has_update, remote_version, local_version, None

        except Exception as e:
            local_version = self.config.current_version
            return False, None, local_version, f"检查更新失败: {str(e)}"

    def download_update(self, version: str, progress_callback=None) -> tuple:
        """
        下载更新文件
        :param version: 要下载的版本号
        :param progress_callback: 进度回调函数
        :return: (是否成功, 下载文件路径, 错误信息)
        """
        try:
            # 获取下载链接
            download_url = self.github_client.get_download_url(version)
            if not download_url:
                return False, None, "无法获取下载链接"

            # 创建备份
            backup_path = self.backup_manager.create_backup()
            if not backup_path:
                return False, None, "创建备份失败"

            # 下载文件
            downloaded_file = self.download_manager.download_file(
                download_url,
                version,
                progress_callback
            )

            if downloaded_file:
                return True, downloaded_file, None
            else:
                return False, None, "下载失败"

        except Exception as e:
            return False, None, f"下载更新失败: {str(e)}"

    def execute_update(self, update_file_path: str, new_version: str) -> tuple:
        """
        执行应用程序更新操作

        Args:
            update_file_path (str): 下载的更新文件完整路径
            new_version (str): 目标版本号，必须符合语义化版本格式 (如 "1.2.3")

        Returns:
            tuple[bool, Optional[str]]: (更新是否成功, 错误信息)

        Raises:
            ValueError: 当版本号格式无效时
            UpdateExecutionError: 当更新执行过程中出现错误时

        Note:
            此方法会自动创建备份并执行文件替换操作
            更新成功后会更新本地版本信息
        """
        try:
            # 参数验证
            if not update_file_path or not update_file_path.strip():
                return False, "更新文件路径不能为空"

            if not new_version or not new_version.strip():
                return False, "新版本号不能为空"

            if not self._is_valid_version_format(new_version):
                return False, f"版本号格式无效: {new_version}"

            # 检查更新文件是否存在
            import os
            if not os.path.exists(update_file_path):
                return False, f"更新文件不存在: {update_file_path}"

            # 执行更新
            success = self.update_executor.execute_update(update_file_path, new_version)
            if success:
                return True, None
            else:
                return False, "更新执行失败"

        except UpdateExecutionError as e:
            # 保留具体的执行错误信息
            return False, f"更新执行失败: {str(e)}"
        except ValueError as e:
            # 参数验证错误处理
            return False, f"参数验证失败: {str(e)}"
        except Exception as e:
            return False, f"执行更新异常: {str(e)}"

    def rollback_update(self) -> tuple:
        """
        回滚更新
        :return: (是否成功, 错误信息)
        """
        try:
            success = self.backup_manager.restore_from_backup()
            if success:
                return True, None
            else:
                return False, "回滚失败"

        except Exception as e:
            return False, f"回滚异常: {str(e)}"

    def _is_valid_version_format(self, version: str) -> bool:
        """
        验证版本号格式是否有效

        Args:
            version (str): 版本号字符串

        Returns:
            bool: 版本号格式是否有效
        """
        try:
            from packaging import version as pkg_version
            pkg_version.parse(version)
            return True
        except Exception:
            return False

    # UI集成接口方法
    def set_parent_window(self, parent):
        """
        设置父窗口引用

        Args:
            parent: 父窗口对象
        """
        self.parent = parent
        if self.ui_manager:
            self.ui_manager.set_parent(parent)

    def show_update_dialog(self) -> None:
        """
        显示更新对话框 - 完整的更新流程
        这是主程序调用的主要接口
        """
        if not self.ui_manager:
            print("UI管理器未初始化")
            return

        try:
            # 检查网络连接
            self._check_network_connection()

            # 检查更新
            has_update, remote_version, local_version, error_msg = self.check_for_updates(force_check=True)

            # 显示更新通知
            user_confirm, error = self.ui_manager.show_update_notification(
                has_update, remote_version, error_msg
            )

            # 如果有更新且用户确认，开始下载
            if has_update and user_confirm:
                self._handle_download_flow(remote_version)

        except Exception as e:
            if self.ui_manager:
                error_msg = str(e).lower()
                if any(keyword in error_msg for keyword in ["network", "连接", "connection"]):
                    self.ui_manager.show_error_dialog("网络错误", f"网络连接出现问题:\n{str(e)}", "warning")
                elif any(keyword in error_msg for keyword in ["timeout", "超时"]):
                    self.ui_manager.show_error_dialog("连接超时", "服务器响应超时，请稍后重试。", "warning")
                else:
                    self.ui_manager.show_error_dialog("更新检查失败", f"检查更新时发生未知错误:\n{str(e)}", "error")
            else:
                print(f"更新流程异常: {e}")

    def handle_startup_check(self) -> None:
        """
        处理启动时的更新检查
        非阻塞方式，不影响应用启动
        """
        if not self.ui_manager:
            return

        try:
            has_update, remote_version, local_version, error_msg = self.check_for_updates()

            if error_msg and "距离上次检查时间过短" not in error_msg:
                print(f"更新检查失败: {error_msg}")
            elif has_update:
                # 延迟显示更新通知，避免干扰应用启动
                from PyQt5.QtCore import QTimer
                QTimer.singleShot(2000, lambda: self._delayed_update_notification(remote_version))

        except Exception as e:
            print(f"启动更新检查异常: {e}")

    def _check_network_connection(self) -> None:
        """检查网络连接"""
        import socket
        try:
            socket.create_connection(("www.github.com", 80), timeout=5)
        except (socket.timeout, socket.error, OSError) as e:
            if self.ui_manager:
                self.ui_manager.show_error_dialog("网络错误", f"无法连接到更新服务器 ({str(e)})", "warning")
            raise

    def _delayed_update_notification(self, remote_version: str) -> None:
        """延迟的更新通知"""
        if self.ui_manager:
            user_confirm, _ = self.ui_manager.show_update_notification(True, remote_version)
            if user_confirm:
                self._handle_download_flow(remote_version)

    def _handle_download_flow(self, version: str) -> None:
        """处理下载流程"""
        try:
            # 获取文件大小信息
            release_info = self.github_client.get_release_info(version)
            file_size = None
            if release_info and 'assets' in release_info and release_info['assets']:
                size_bytes = release_info['assets'][0].get('size', 0)
                if self.ui_manager:
                    file_size = self.ui_manager.format_file_size(size_bytes)

            # 显示下载确认
            if not self.ui_manager.show_download_confirm(version, file_size):
                return

            # 创建进度对话框
            progress_dialog = self.ui_manager.create_progress_dialog("下载更新")

            # 定义进度回调
            def progress_callback(downloaded: int, total: int, percentage: float):
                if self.ui_manager and progress_dialog:
                    self.ui_manager.update_progress(progress_dialog, downloaded, total, percentage)

            # 开始下载
            success, download_path, error = self.download_update(version, progress_callback)

            # 关闭进度对话框
            if self.ui_manager:
                self.ui_manager.close_progress_dialog(progress_dialog)

            if success and download_path:
                # 下载成功，询问是否安装
                if self.ui_manager.show_install_confirm(version):
                    # 执行更新
                    update_success, update_error = self.execute_update(download_path, version)

                    if update_success:
                        self.ui_manager.show_update_complete(version)
                    else:
                        self.ui_manager.show_error_dialog("更新失败", update_error, "error")
            else:
                # 下载失败
                if self.ui_manager:
                    self.ui_manager.show_error_dialog("下载失败", error or "未知错误", "error")

        except Exception as e:
            if self.ui_manager:
                self.ui_manager.show_error_dialog("下载流程异常", str(e), "error")
            else:
                print(f"下载流程异常: {e}")

    def show_update_dialog_async(self) -> None:
        """
        显示异步更新对话框 - 完整的异步更新流程
        这是主程序调用的主要接口，支持真正的异步下载
        """
        logger.info("开始异步更新流程")

        if not self.ui_manager:
            logger.error("UI管理器未初始化")
            print("UI管理器未初始化")
            return

        try:
            # 检查网络连接
            self._check_network_connection()

            # 检查更新
            has_update, remote_version, local_version, error_msg = self.check_for_updates(force_check=True)

            logger.info(f"更新检查结果: has_update={has_update}, remote_version={remote_version}, local_version={local_version}")

            # 显示更新通知
            user_confirm, error = self.ui_manager.show_update_notification(
                has_update, remote_version, error_msg
            )

            logger.info(f"用户选择: {user_confirm}")

            # 如果有更新且用户确认，开始异步下载
            if has_update and user_confirm:
                logger.info(f"开始异步下载版本 {remote_version}")
                self._handle_async_download_flow(remote_version)
            elif has_update and not user_confirm:
                logger.info("用户取消下载")

        except Exception as e:
            logger.error(f"异步更新流程异常: {e}", exc_info=True)
            if self.ui_manager:
                error_msg = str(e).lower()
                if any(keyword in error_msg for keyword in ["network", "连接", "connection"]):
                    self.ui_manager.show_error_dialog("网络错误", f"网络连接出现问题:\n{str(e)}", "warning")
                elif any(keyword in error_msg for keyword in ["timeout", "超时"]):
                    self.ui_manager.show_error_dialog("连接超时", "服务器响应超时，请稍后重试。", "warning")
                else:
                    self.ui_manager.show_error_dialog("更新检查失败", f"检查更新时发生未知错误:\n{str(e)}", "error")
            else:
                print(f"异步更新流程异常: {e}")

    def _handle_async_download_flow(self, version: str) -> None:
        """处理异步下载流程"""
        try:
            logger.info(f"开始处理异步下载流程，版本: {version}")

            # 获取下载链接
            try:
                download_url = self.github_client.get_download_url(version)
                if not download_url:
                    if self.ui_manager:
                        self.ui_manager.show_error_dialog(ERROR_DOWNLOAD_URL_TITLE, ERROR_DOWNLOAD_URL_FAILED, "error")
                    return
            except (NetworkError, VersionCheckError) as e:
                logger.error(f"获取下载链接异常: {str(e)}")
                if self.ui_manager:
                    self.ui_manager.show_error_dialog(ERROR_DOWNLOAD_URL_TITLE, str(e), "critical")
                return
            except Exception as e:
                logger.error(f"获取下载链接未知异常: {str(e)}")
                if self.ui_manager:
                    self.ui_manager.show_error_dialog(ERROR_DOWNLOAD_URL_TITLE, f"未知错误: {str(e)}", "critical")
                return

            logger.info(f"获取到下载链接: {download_url}")

            # 检查下载状态，防止重复调用
            if self._download_state == DownloadState.DOWNLOADING and self._current_download_version == version:
                logger.warning(f"版本 {version} 正在下载中，跳过重复请求")
                return

            # 显示下载确认对话框（异步获取文件大小，一次确认）
            confirmed = self._show_download_confirm_async(version, download_url)
            if not confirmed:
                logger.info("用户取消下载确认")
                return

            # 设置下载状态
            self._download_state = DownloadState.DOWNLOADING
            self._current_download_version = version

            try:
                # 创建异步下载任务
                download_thread, file_path = self.download_manager.download_file_async(download_url, version)

                logger.info(f"创建异步下载任务: {file_path}")

                # 使用UI管理器进行异步下载
                success, download_path, error = self.ui_manager.download_with_dialog_async(
                    download_thread, version
                )

                logger.info(f"异步下载结果: success={success}, download_path={download_path}, error={error}")

                # 更新下载状态
                if success:
                    self._download_state = DownloadState.COMPLETED
                    logger.info(f"版本 {version} 下载完成")
                else:
                    self._download_state = DownloadState.FAILED
                    logger.error(f"版本 {version} 下载失败: {error}")

                # 继续处理下载结果（这部分逻辑已经在try块中）
                if success and download_path:
                    # 验证下载的文件
                    try:
                        if self.download_manager.verify_downloaded_file(download_path):
                            logger.info("文件验证成功")
                            # 下载成功，询问是否安装
                            if self.ui_manager.show_install_confirm(version):
                                logger.info(f"用户确认安装版本 {version}")
                                # 执行更新
                                update_success, update_error = self.execute_update(download_path, version)

                                if update_success:
                                    logger.info("更新执行成功")
                                    self.ui_manager.show_update_complete(version)
                                else:
                                    logger.error(f"更新执行失败: {update_error}")
                                    self.ui_manager.show_error_dialog("更新失败", update_error, "error")
                            else:
                                logger.info("用户取消安装")
                        else:
                            error_msg = "下载文件验证失败，文件可能损坏"
                            logger.error(error_msg)
                            if self.ui_manager:
                                self.ui_manager.show_error_dialog("文件验证失败", error_msg, "error")

                    except Exception as e:
                        error_msg = f"文件验证异常: {str(e)}"
                        logger.error(error_msg)
                        if self.ui_manager:
                            self.ui_manager.show_error_dialog("文件验证异常", error_msg, "error")

            except Exception as e:
                self._download_state = DownloadState.FAILED
                logger.error(f"版本 {version} 下载过程异常: {str(e)}", exc_info=True)
                if self.ui_manager:
                    error_msg = f"下载过程中发生异常: {str(e)}"
                    self.ui_manager.show_error_dialog("下载异常", error_msg, "critical")
            finally:
                # 重置下载状态（延迟重置，防止快速重复调用）
                import threading
                def reset_state():
                    import time
                    time.sleep(2)  # 延迟2秒重置状态
                    if self._download_state == DownloadState.COMPLETED or self._download_state == DownloadState.FAILED:
                        self._download_state = DownloadState.IDLE
                        self._current_download_version = None

                reset_thread = threading.Thread(target=reset_state, daemon=True)
                reset_thread.start()

        except Exception as e:
            logger.error(f"异步下载流程异常: {e}", exc_info=True)
            if self.ui_manager:
                self.ui_manager.show_error_dialog("下载流程异常", str(e), "error")
            else:
                print(f"异步下载流程异常: {e}")

    def download_update_async(self, version: str, progress_callback: Optional[Callable] = None) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        异步下载更新文件

        Args:
            version: 版本号
            progress_callback: 进度回调函数（可选）

        Returns:
            tuple: (是否成功, 文件路径, 错误信息)
        """
        try:
            # 获取下载链接
            download_url, error = self.github_client.get_download_url(version)
            if not download_url:
                return False, None, error or "获取下载链接失败"

            # 创建异步下载任务
            download_thread, file_path = self.download_manager.download_file_async(download_url, version)

            if self.ui_manager and progress_callback:
                # 使用UI管理器进行异步下载，支持进度回调
                return self.ui_manager.download_with_dialog_async(download_thread, version, progress_callback)
            else:
                # 直接启动下载线程（无UI）
                download_thread.start()
                download_thread.wait()  # 等待完成
                # 注意：这里简化了处理，实际应用中可能需要更复杂的逻辑
                return True, file_path, None

        except Exception as e:
            logger.error(f"异步下载更新失败: {e}", exc_info=True)
            error_msg = f"异步下载异常: {type(e).__name__}: {str(e)}"
            return False, None, error_msg

    def is_async_supported(self) -> bool:
        """
        检查是否支持异步下载

        Returns:
            bool: 是否支持异步下载
        """
        try:
            from .ui.async_download_thread import AsyncDownloadThread
            return True
        except ImportError:
            return False

    def _show_download_confirm_async(self, version: str, download_url: str) -> bool:
        """
        异步显示下载确认对话框（一次确认，动态更新文件信息）
        :param version: 版本号
        :param download_url: 下载链接
        :return: 用户是否确认下载
        """
        import threading

        file_size = None
        size_thread = None

        def get_file_size_async():
            """异步获取文件大小"""
            try:
                nonlocal file_size
                file_size = self.download_manager.get_download_size_cached(download_url)
            except Exception as e:
                logger.error(f"异步获取文件大小失败: {str(e)}")

        # 启动异步线程获取文件大小
        size_thread = threading.Thread(target=get_file_size_async, daemon=True)
        size_thread.start()

        # 显示确认对话框（包含加载状态）
        if self.ui_manager:
            confirmed = self.ui_manager.show_download_confirm(
                version,
                "正在获取文件信息...",
                show_loading=True
            )

            # 等待文件大小获取完成（使用配置化超时）
            if confirmed and size_thread:
                size_thread.join(timeout=NETWORK_TIMEOUT_MEDIUM * 2)

                # 如果获取到了文件大小，显示更新后的信息
                if file_size:
                    file_size_str = self.ui_manager.format_file_size(file_size)
                    logger.info(f"获取到文件大小: {file_size_str}")
                else:
                    logger.warning("未能获取到文件大小，继续执行")

            return confirmed

        return False

# 导出的公共接口
__all__ = [
    'AutoUpdater',
    'GitHubClient',
    'DownloadManager',
    'BackupManager',
    'UpdateExecutor',
    'UpdateError',
    'NetworkError',
    'VersionCheckError',
    'DownloadError',
    'BackupError',
    'UpdateExecutionError'
]