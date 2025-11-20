"""
Update UI Manager

Central UI controller for update operations. Manages all user interface
interactions during the update process including dialogs, progress displays,
and user notifications.
"""

from typing import Optional, Callable, Any
from PyQt5.QtWidgets import QMessageBox, QWidget
from PyQt5.QtCore import QTimer
import logging

from .update_dialogs import UpdateDialogs
from .progress_dialog import ProgressDialog

logger = logging.getLogger(__name__)


class UpdateUIManager:
    """UI管理器，负责更新过程中的所有用户交互"""

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化UI管理器

        Args:
            parent: 父窗口对象，用于居中显示对话框
        """
        self.parent = parent
        self.update_dialogs = UpdateDialogs(parent)
        self.progress_dialog = ProgressDialog(parent)

    def set_parent(self, parent: QWidget):
        """设置父窗口"""
        self.parent = parent
        self.update_dialogs.set_parent(parent)
        self.progress_dialog.set_parent(parent)

    def show_update_notification(self, has_update: bool, remote_version: str = None,
                                error_msg: str = None) -> tuple:
        """
        显示更新通知

        Args:
            has_update: 是否有更新
            remote_version: 远程版本号
            error_msg: 错误信息

        Returns:
            tuple: (用户选择结果, 错误信息)
        """
        return self.update_dialogs.show_update_notification(has_update, remote_version, error_msg)

    def show_download_confirm(self, version: str, file_size: str = None) -> bool:
        """
        显示下载确认对话框

        Args:
            version: 版本号
            file_size: 文件大小

        Returns:
            bool: 用户是否确认下载
        """
        return self.update_dialogs.show_download_confirm(version, file_size)

    def show_install_confirm(self, version: str) -> bool:
        """
        显示安装确认对话框

        Args:
            version: 版本号

        Returns:
            bool: 用户是否确认安装
        """
        return self.update_dialogs.show_install_confirm(version)

    def show_update_complete(self, version: str, needs_restart: bool = True) -> None:
        """
        显示更新完成对话框

        Args:
            version: 更新到的版本
            needs_restart: 是否需要重启应用
        """
        self.update_dialogs.show_update_complete(version, needs_restart)

    def create_progress_dialog(self, title: str = "下载更新") -> Any:
        """
        创建进度对话框

        Args:
            title: 对话框标题

        Returns:
            进度对话框对象
        """
        return self.progress_dialog.create_progress_dialog(title)

    def update_progress(self, dialog: Any, downloaded: int, total: int,
                       percentage: float, extra_info: str = None) -> None:
        """
        更新进度显示

        Args:
            dialog: 进度对话框对象
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 完成百分比
            extra_info: 额外信息
        """
        self.progress_dialog.update_progress(dialog, downloaded, total, percentage, extra_info)

    def close_progress_dialog(self, dialog: Any) -> None:
        """
        关闭进度对话框

        Args:
            dialog: 进度对话框对象
        """
        self.progress_dialog.close_progress_dialog(dialog)

    def show_error_dialog(self, title: str, message: str, error_type: str = "error") -> None:
        """
        显示错误对话框

        Args:
            title: 对话框标题
            message: 错误信息
            error_type: 错误类型 (warning/error/critical)
        """
        self.update_dialogs.show_error_dialog(title, message, error_type)

    def format_file_size(self, size_bytes: int) -> str:
        """
        格式化文件大小显示

        Args:
            size_bytes: 文件大小（字节）

        Returns:
            str: 格式化后的文件大小
        """
        return self.progress_dialog.format_file_size(size_bytes)