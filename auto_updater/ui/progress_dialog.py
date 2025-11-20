"""
Progress Dialog

Provides progress display functionality for download operations including
file size formatting, progress percentage calculation, and UI updates.
"""

from typing import Optional
from PyQt5.QtWidgets import QMessageBox, QWidget
import logging

logger = logging.getLogger(__name__)


class ProgressDialog:
    """进度对话框管理器"""

    def __init__(self, parent: Optional[QWidget] = None):
        """
        初始化进度对话框管理器

        Args:
            parent: 父窗口对象
        """
        self.parent = parent
        self._last_update_time = 0
        self._update_interval = 0.5  # UI更新间隔（秒）

    def set_parent(self, parent: QWidget):
        """设置父窗口"""
        self.parent = parent

    def create_progress_dialog(self, title: str = "下载更新") -> QMessageBox:
        """
        创建进度对话框

        Args:
            title: 对话框标题

        Returns:
            QMessageBox: 进度对话框对象
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle(title)
            msg_box.setText("准备下载...")
            msg_box.setStandardButtons(QMessageBox.NoButton)
            msg_box.show()

            return msg_box
        except Exception as e:
            logger.error(f"创建进度对话框失败: {e}")
            raise

    def update_progress(self, dialog: QMessageBox, downloaded: int, total: int,
                       percentage: float, extra_info: str = None) -> None:
        """
        更新进度显示（优化版本）

        Args:
            dialog: 进度对话框
            downloaded: 已下载字节数
            total: 总字节数
            percentage: 完成百分比
            extra_info: 额外信息
        """
        try:
            import time

            current_time = time.time()
            # 性能优化：控制更新频率
            if current_time - self._last_update_time < self._update_interval:
                return

            self._last_update_time = current_time

            # 格式化文件大小和进度文本
            downloaded_str = self.format_file_size(downloaded)
            total_str = self.format_file_size(total)

            # 构建进度文本
            if extra_info:
                progress_text = f"{extra_info}\n{percentage:.1f}%\n已下载: {downloaded_str} / {total_str}"
            else:
                progress_text = f"正在下载更新... {percentage:.1f}%\n已下载: {downloaded_str} / {total_str}"

            # 更新对话框文本
            dialog.setText(progress_text)

            # 确保对话框显示在最前面
            dialog.raise_()
            dialog.activateWindow()

        except Exception as e:
            logger.error(f"更新进度显示失败: {e}")

    def close_progress_dialog(self, dialog: Optional[QMessageBox]) -> None:
        """
        关闭进度对话框

        Args:
            dialog: 进度对话框对象
        """
        try:
            if dialog and hasattr(dialog, 'close'):
                dialog.close()
        except Exception as e:
            logger.error(f"关闭进度对话框失败: {e}")

    def format_file_size(self, size_bytes: int) -> str:
        """
        格式化文件大小显示

        Args:
            size_bytes: 文件大小（字节）

        Returns:
            str: 格式化后的文件大小
        """
        try:
            if size_bytes == 0:
                return "0 B"

            # 定义单位
            units = ['B', 'KB', 'MB', 'GB', 'TB']
            unit_index = 0
            size = float(size_bytes)

            # 计算合适的单位
            while size >= 1024 and unit_index < len(units) - 1:
                size /= 1024
                unit_index += 1

            # 格式化输出
            if unit_index == 0:  # 字节
                return f"{int(size)} {units[unit_index]}"
            else:  # KB及以上
                return f"{size:.1f} {units[unit_index]}"

        except Exception as e:
            logger.error(f"格式化文件大小失败: {e}")
            return f"{size_bytes} B"

    def create_status_update_dialog(self, title: str, message: str) -> QMessageBox:
        """
        创建状态更新对话框

        Args:
            title: 对话框标题
            message: 显示信息

        Returns:
            QMessageBox: 状态对话框对象
        """
        try:
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle(title)
            msg_box.setText(message)
            msg_box.setStandardButtons(QMessageBox.NoButton)
            msg_box.show()

            return msg_box
        except Exception as e:
            logger.error(f"创建状态更新对话框失败: {e}")
            raise

    def update_status_message(self, dialog: QMessageBox, message: str) -> None:
        """
        更新状态信息

        Args:
            dialog: 对话框对象
            message: 新的状态信息
        """
        try:
            dialog.setText(message)
            dialog.raise_()
            dialog.activateWindow()
        except Exception as e:
            logger.error(f"更新状态信息失败: {e}")