import os
from tkinter import messagebox
from tkinterdnd2 import DND_FILES


def setup_file_drop(
    widget,
    callback,
    extensions=(".xlsx", ".xlsm"),
    normal_text="📥 KÉO THẢ FILE VÀO ĐÂY",
    hover_text="📥 THẢ FILE TẠI ĐÂY",
):
    """
    Đăng ký vùng kéo-thả file.

    widget:
        Widget Tkinter nhận file.

    callback:
        Hàm được gọi khi file hợp lệ được thả.
        callback(file_path)

    extensions:
        Các phần mở rộng được phép.
    """

    widget.drop_target_register(DND_FILES)

    def on_drop(event):
        try:
            files = widget.tk.splitlist(event.data)

            if not files:
                return

            file_path = files[0].strip().strip('"')

            if not file_path.lower().endswith(extensions):
                messagebox.showerror(
                    "File không hợp lệ", "Định dạng file không được hỗ trợ."
                )
                return

            if not os.path.isfile(file_path):
                messagebox.showerror(
                    "Không tìm thấy file", f"Không tìm thấy:\n\n{file_path}"
                )
                return

            callback(file_path)

        except Exception as exc:
            messagebox.showerror("Lỗi kéo thả file", f"Không thể xử lý file:\n\n{exc}")

    widget.dnd_bind("<<Drop>>", on_drop)

    widget.dnd_bind(
        "<<DragEnter>>",
        lambda event: widget.config(
            relief="sunken",
            text=hover_text,
        ),
    )

    widget.dnd_bind(
        "<<DragLeave>>",
        lambda event: widget.config(
            relief="groove",
            text=normal_text,
        ),
    )
