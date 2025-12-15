# -*- coding: utf-8 -*-
"""
自动更新器配置常量
将JSON配置信息转换为Python常量，消除外部文件依赖
"""

# 应用配置
APP_NAME: str = "SAP操作工具"
APP_EXECUTABLE: str = "Sap_Operate_theme.exe"
APP_EXECUTABLE_DEV: str = "Sap_Operate_theme.py"  # 开发环境可执行文件
APP_EXECUTABLE_PROD: str = "Sap_Operate_theme.exe"     # 生产环境可执行文件

# GitHub仓库配置
GITHUB_OWNER: str = "chen-huai"
GITHUB_REPO: str = "Sap_Operation"
GITHUB_API_BASE: str = "https://api.github.com"

# 版本配置
CURRENT_VERSION: str = "0.1.0"
UPDATE_CHECK_INTERVAL_DAYS: int = 30

# 错误消息常量
ERROR_DOWNLOAD_URL_FAILED: str = "获取下载链接失败，请检查网络连接或版本号"
ERROR_DOWNLOAD_URL_TITLE: str = "获取下载链接失败"

# 网络请求优化配置
NETWORK_TIMEOUT_SHORT: int = 10    # 短超时：HEAD请求、连接测试
NETWORK_TIMEOUT_MEDIUM: int = 20   # 中超时：文件大小获取
NETWORK_TIMEOUT_LONG: int = 60     # 长超时：文件下载
NETWORK_MAX_RETRIES: int = 3       # 最大重试次数
NETWORK_RETRY_DELAY: float = 1.0   # 重试基础延迟（秒）

# 文件大小缓存配置
FILE_SIZE_CACHE_TTL: int = 300     # 缓存有效期（秒）
AUTO_CHECK_ENABLED: bool = True
SHOW_VERSION_IN_FILENAME: bool = False  # 控制下载文件名是否包含版本号

# 更新配置
MAX_BACKUP_COUNT: int = 3
DOWNLOAD_TIMEOUT: int = 300
MAX_RETRIES: int = 3
AUTO_RESTART: bool = True

# 网络配置
REQUEST_HEADERS: dict = {
    "Accept": "application/vnd.github.v3+json",
    "User-Agent": "Sap_Operation/1.0"
}

# 便利常量（兼容现有API）
GITHUB_REPO_PATH: str = f"{GITHUB_OWNER}/{GITHUB_REPO}"
GITHUB_RELEASES_URL: str = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO_PATH}/releases"
GITHUB_LATEST_RELEASE_URL: str = f"{GITHUB_RELEASES_URL}/latest"

# 默认配置字典（保持JSON格式兼容）
DEFAULT_CONFIG: dict = {
    "app": {
        "name": APP_NAME,
        "executable": APP_EXECUTABLE
    },
    "repository": {
        "owner": GITHUB_OWNER,
        "repo": GITHUB_REPO,
        "api_base": GITHUB_API_BASE
    },
    "version": {
        "current": CURRENT_VERSION,
        "check_interval_days": UPDATE_CHECK_INTERVAL_DAYS,
        "auto_check_enabled": AUTO_CHECK_ENABLED
    },
    "update": {
        "backup_count": MAX_BACKUP_COUNT,
        "download_timeout": DOWNLOAD_TIMEOUT,
        "max_retries": MAX_RETRIES,
        "auto_restart": AUTO_RESTART
    },
    "network": {
        "request_headers": REQUEST_HEADERS
    }
}

# 版本信息验证
def validate_version_format(version_str: str) -> bool:
    """验证版本号格式是否有效"""
    try:
        from packaging import version as pkg_version
        pkg_version.parse(version_str)
        return True
    except Exception:
        return False

# 配置完整性验证
def validate_config() -> bool:
    """验证配置信息的完整性"""
    try:
        # 验证必要常量
        required_constants = [
            APP_NAME, APP_EXECUTABLE, GITHUB_OWNER, GITHUB_REPO,
            GITHUB_API_BASE, CURRENT_VERSION, REQUEST_HEADERS
        ]

        for const in required_constants:
            if not const:
                return False

        # 验证版本号格式
        if not validate_version_format(CURRENT_VERSION):
            return False

        # 验证URL格式
        if not GITHUB_API_BASE.startswith("https://"):
            return False

        return True
    except Exception:
        return False

# 在模块加载时验证配置
if not validate_config():
    raise ValueError("配置信息验证失败，请检查config_constants.py中的配置")