import os
import subprocess
import sys
from tkinter import messagebox


def open_file(path):
    path = os.path.abspath(path)

    if not os.path.exists(path):
        return

    try:
        if sys.platform.startswith("win"):
            os.startfile(path)

        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])

        else:
            subprocess.Popen(["xdg-open", path])

    except Exception as exc:  # noqa: BLE001
        messagebox.showwarning(
            "Không mở được file",
            f"File đã được tạo nhưng không thể tự mở:\n{exc}",
        )
