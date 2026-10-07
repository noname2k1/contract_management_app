# ============================================================
# SCROLLABLE FRAME
# ============================================================
# Không còn canvas lồng với main.py.
# Chỉ màn hình này quản lý scrollbar.
# ============================================================


from tkinter import ttk
from tkinter import messagebox
import tkinter as tk


class ScrollableFrame(ttk.Frame):
    def __init__(
        self,
        parent,
    ):
        super().__init__(parent)

        # ========================================================
        # CANVAS
        # ========================================================

        self.canvas = tk.Canvas(
            self,
            highlightthickness=0,
            borderwidth=0,
        )

        # ========================================================
        # SCROLLBAR DỌC
        # ========================================================

        self.v_scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.canvas.yview,
        )

        # ========================================================
        # SCROLLBAR NGANG
        # ========================================================

        self.h_scrollbar = ttk.Scrollbar(
            self,
            orient="horizontal",
            command=self.canvas.xview,
        )

        # ========================================================
        # FRAME CHỨA TOÀN BỘ NỘI DUNG
        # ========================================================

        self.inner = ttk.Frame(self.canvas)

        self.window_id = self.canvas.create_window(
            (0, 0),
            window=self.inner,
            anchor="nw",
        )

        # ========================================================
        # KẾT NỐI SCROLLBAR
        # ========================================================

        self.canvas.configure(
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
        )

        # ========================================================
        # LAYOUT
        # ========================================================

        self.canvas.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        self.v_scrollbar.grid(
            row=0,
            column=1,
            sticky="ns",
        )

        self.h_scrollbar.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        self.rowconfigure(
            0,
            weight=1,
        )

        self.columnconfigure(
            0,
            weight=1,
        )

        # ========================================================
        # CẬP NHẬT VÙNG SCROLL
        # ========================================================

        self.inner.bind(
            "<Configure>",
            self._update_scrollregion,
        )

        self.canvas.bind(
            "<Configure>",
            self._on_canvas_configure,
        )

        # ========================================================
        # MOUSE WHEEL
        # ========================================================

        self.canvas.bind(
            "<Enter>",
            self._bind_mousewheel,
        )

        self.canvas.bind(
            "<Leave>",
            self._unbind_mousewheel,
        )

        self.inner.bind(
            "<Enter>",
            self._bind_mousewheel,
        )

        self.inner.bind(
            "<Leave>",
            self._unbind_mousewheel,
        )

        # ========================================================
        # SHIFT + MOUSE WHEEL
        # -> CUỘN NGANG
        # ========================================================

        self.canvas.bind(
            "<Shift-MouseWheel>",
            self._on_shift_mousewheel,
        )

        self.inner.bind(
            "<Shift-MouseWheel>",
            self._on_shift_mousewheel,
        )

        # ========================================================
        # PHÍM MŨI TÊN
        # ========================================================

        self.canvas.bind(
            "<Left>",
            self._scroll_left,
        )

        self.canvas.bind(
            "<Right>",
            self._scroll_right,
        )

        # ========================================================
        # LUÔN BẮT ĐẦU Ở ĐẦU
        # ========================================================

        self.after_idle(self.scroll_to_top)

    # ========================================================
    # SCROLL REGION
    # ========================================================

    def _update_scrollregion(
        self,
        event=None,
    ):
        bbox = self.canvas.bbox("all")

        if bbox:
            self.canvas.configure(scrollregion=bbox)

    # ========================================================
    # CANVAS RESIZE
    # ========================================================

    def _on_canvas_configure(
        self,
        event,
    ):
        """
        Không ép chiều rộng inner bằng canvas.

        Đây là điểm quan trọng để scrollbar ngang hoạt động.
        """

        self._update_scrollregion()

    # ========================================================
    # BIND MOUSE WHEEL
    # ========================================================

    def _bind_mousewheel(
        self,
        event=None,
    ):
        self.canvas.bind_all(
            "<MouseWheel>",
            self._on_mousewheel,
        )

    # ========================================================
    # UNBIND MOUSE WHEEL
    # ========================================================

    def _unbind_mousewheel(
        self,
        event=None,
    ):
        self.canvas.unbind_all("<MouseWheel>")

    # ========================================================
    # CUỘN DỌC
    # ========================================================

    def _on_mousewheel(
        self,
        event,
    ):
        if event.delta:
            self.canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units",
            )

    # ========================================================
    # SHIFT + MOUSE WHEEL
    # -> CUỘN NGANG
    # ========================================================

    def _on_shift_mousewheel(
        self,
        event,
    ):
        if event.delta:
            self.canvas.xview_scroll(
                int(-1 * (event.delta / 120)),
                "units",
            )

        return "break"

    # ========================================================
    # CUỘN SANG TRÁI
    # ========================================================

    def _scroll_left(
        self,
        event=None,
    ):
        self.canvas.xview_scroll(
            -3,
            "units",
        )

        return "break"

    # ========================================================
    # CUỘN SANG PHẢI
    # ========================================================

    def _scroll_right(
        self,
        event=None,
    ):
        self.canvas.xview_scroll(
            3,
            "units",
        )

        return "break"

    # ========================================================
    # VỀ ĐẦU
    # ========================================================

    def scroll_to_top(
        self,
    ):
        self.canvas.yview_moveto(0)

        self.canvas.xview_moveto(0)
