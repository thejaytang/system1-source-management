from __future__ import annotations

import os
import platform
import subprocess


def notify(title: str, message: str) -> None:
    """Best-effort native notification; logs remain the authoritative record."""
    if os.environ.get("SYSTEM1_DISABLE_NOTIFICATIONS") == "1":
        return
    system = platform.system()
    try:
        if system == "Darwin":
            safe_title = title.replace('"', "'")
            safe_message = message.replace('"', "'")
            subprocess.run(["osascript", "-e", f'display notification "{safe_message}" with title "{safe_title}"'], check=False, timeout=10)
        elif system == "Windows":
            script = (
                "$ws=New-Object -ComObject WScript.Shell;"
                f"$ws.Popup('{message.replace("'", "''")}',10,'{title.replace("'", "''")}',64)"
            )
            subprocess.run(["powershell", "-NoProfile", "-Command", script], check=False, timeout=15)
    except (OSError, subprocess.SubprocessError):
        return

