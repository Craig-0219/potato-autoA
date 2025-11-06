"""AUTOA RPA core package."""
import sys
from pathlib import Path

__all__ = ["ui", "line_automation", "resource_path"]


def resource_path(relative_path: str | Path) -> Path:
    """
    取得資源檔案的絕對路徑。

    在開發環境中，返回相對於專案根目錄的路徑。
    在 PyInstaller 打包後，返回臨時解壓目錄中的路徑。

    Args:
        relative_path: 相對路徑 (例如: "templates/friend.png")

    Returns:
        資源檔案的絕對路徑
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包後的臨時目錄
        base_path = Path(sys._MEIPASS)
    else:
        # 開發環境：使用專案根目錄
        base_path = Path(__file__).parent.parent

    return base_path / relative_path
