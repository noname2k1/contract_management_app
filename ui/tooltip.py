import tkinter as tk


class ToolTip:
    """Tooltip nhỏ hiển thị khi rê chuột lên nút."""

    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self.after_id = None

        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self.after_id = self.widget.after(
            self.delay,
            self._show,
        )

    def _cancel(self):
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:  # noqa: BLE001, S110
                pass
            self.after_id = None

    def _show(self):
        if self.tipwindow is not None:
            return

        try:
            x = self.widget.winfo_rootx() + self.widget.winfo_width() + 8
            y = self.widget.winfo_rooty() + max(0, self.widget.winfo_height() // 2 - 12)

            self.tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")

            label = tk.Label(
                tw,
                text=self.text,
                justify="left",
                background="#333333",
                foreground="white",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=5,
                font=("Arial", 9),
            )
            label.pack()
        except Exception:
            self.tipwindow = None

    def _hide(self, _event=None):
        self._cancel()
        if self.tipwindow is not None:
            try:
                self.tipwindow.destroy()
            except Exception:
                pass
            self.tipwindow = None
