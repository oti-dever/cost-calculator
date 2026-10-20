"""桌面界面共用的系统操作。"""

import os
import subprocess
import sys
from pathlib import Path


def open_with_default_app(path: str | Path) -> None:
    """使用操作系统默认应用打开文件。"""

    resolved = str(Path(path).resolve())
    if hasattr(os, "startfile"):
        os.startfile(resolved)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", resolved])
    else:
        subprocess.Popen(["xdg-open", resolved])
