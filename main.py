import os
from tkinter import messagebox, ttk

from configs import style_config
from hop_dong_duyet_gia import HopDongDuyetGiaFrame
from nghiem_thu_thanh_ly import NghiemThuFrame
from tkinterdnd2 import TkinterDnD
from ui import ToolTip


class MainApp:
    def __init__(self, root):
        self.root = root

        self.root.title("Hệ thống quản lý tài liệu")
        self.root.geometry("1200x750")
        self.root.minsize(900, 550)

        # Trạng thái sidebar
        self.sidebar_open_width = 235
        self.sidebar_closed_width = 58
        self.sidebar_buttons = []

        # ==================================================
        # STYLE DÙNG CHUNG TOÀN BỘ ỨNG DỤNG
        # ==================================================
        style_config()
        # ==================================================
        # HEADER
        # ==================================================
        header = ttk.Frame(self.root)
        header.pack(
            fill="x",
            padx=10,
            pady=10,
        )
        ttk.Label(
            header,
            text="HỆ THỐNG QUẢN LÝ TÀI LIỆU",
            style="AppTitle.TLabel",
        ).pack(side="left")
        # ==================================================
        # BODY
        # ==================================================
        body = ttk.Frame(self.root)
        body.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=10,
        )
        # ==================================================
        # SIDEBAR
        # ==================================================

        self.sidebar = ttk.LabelFrame(
            body,
            text="Chức năng",
        )

        self.sidebar.pack(
            side="left",
            fill="y",
            padx=(0, 15),
        )

        self.sidebar.configure(width=self.sidebar_open_width)
        self.sidebar.pack_propagate(False)

        # Nút thu gọn / mở rộng
        self.sidebar_toggle = ttk.Button(
            self.sidebar,
            text="◀◀",
            style="Menu.TButton",
            command=self.toggle_sidebar,
        )
        self.sidebar_toggle.pack(
            fill="x",
            padx=6,
            pady=(8, 10),
        )
        ToolTip(
            self.sidebar_toggle,
            "Thu gọn / mở rộng thanh chức năng",
        )

        # Các nút menu
        self._create_sidebar_button(
            "📄",
            "Tạo duyệt giá - hợp đồng",
            self.show_documents,
        )

        self._create_sidebar_button(
            "📋",
            "Biên bản nghiệm thu",
            self.show_nghiem_thu,
        )

        self._create_sidebar_button(
            "⚙",
            "Cài đặt",
            self.show_settings,
        )

        self._create_sidebar_button(
            "📁",
            "Mở thư mục templates",
            lambda: self.open_dir("./templates"),
        )

        self._create_sidebar_button(
            "📁",
            "Mở thư mục outputs",
            lambda: self.open_dir("./outputs"),
        )

        self._create_sidebar_button(
            "🗑",
            "Xóa file outputs",
            self.clear_outputs,
        )

        self._create_sidebar_button(
            "❌",
            "Thoát",
            self.exit_app,
        )

        # Mặc định thu nhỏ sidebar
        self.sidebar_collapsed = False
        self.toggle_sidebar()
        # ==================================================
        # CONTENT
        # ==================================================

        content_container = ttk.Frame(body)

        content_container.pack(
            side="left",
            fill="both",
            expand=True,
        )

        self.content = ttk.Frame(
            content_container,
            padding=5,
        )

        self.content.pack(
            fill="both",
            expand=True,
        )

        # ==================================================
        # HOME
        # ==================================================

        self.show_home()

    # ======================================================
    # SIDEBAR
    # ======================================================

    def _create_sidebar_button(self, icon, label, command):
        """Tạo nút sidebar và lưu lại để đổi giữa 2 trạng thái."""
        button = ttk.Button(
            self.sidebar,
            text=f"{icon}  {label}",
            style="Menu.TButton",
            command=command,
        )
        button.pack(
            fill="x",
            padx=6,
            pady=5,
        )

        ToolTip(button, label)

        self.sidebar_buttons.append(
            {
                "button": button,
                "icon": icon,
                "label": label,
            }
        )

        return button

    def toggle_sidebar(self):
        """Thu gọn sidebar thành một cột icon hoặc mở lại."""
        self.sidebar_collapsed = not self.sidebar_collapsed

        if self.sidebar_collapsed:
            self.sidebar.configure(
                width=self.sidebar_closed_width,
                text="",
            )

            self.sidebar_toggle.configure(
                text="▶▶",
            )

            for item in self.sidebar_buttons:
                item["button"].configure(
                    text=item["icon"],
                    width=3,
                )

        else:
            self.sidebar.configure(
                width=self.sidebar_open_width,
                text="Chức năng",
            )

            self.sidebar_toggle.configure(
                text="◀◀",
            )

            for item in self.sidebar_buttons:
                item["button"].configure(
                    text=f"{item['icon']}  {item['label']}",
                    width=0,
                )

        # Cập nhật giao diện ngay lập tức
        self.root.update_idletasks()

    # ======================================================
    # CLEAR CONTENT
    # ======================================================

    def clear_content(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    # ======================================================
    # HOME
    # ======================================================

    def show_home(self):
        self.clear_content()

        frame = ttk.Frame(self.content)

        frame.pack(
            expand=True,
        )

        ttk.Label(
            frame,
            text="Chào mừng bạn!",
            font=("Arial", 24, "bold"),
        ).pack(pady=20)

        ttk.Label(
            frame,
            text="Vui lòng chọn một chức năng ở menu bên trái.",
            font=("Arial", 13),
        ).pack()

    # ======================================================
    # DOCUMENTS
    # ======================================================

    def show_documents(self):
        self.clear_content()

        frame = HopDongDuyetGiaFrame(self.content)

        frame.pack(
            fill="both",
            expand=True,
        )

        # Đảm bảo màn hình luôn bắt đầu từ đầu.
        if hasattr(frame, "scroll_to_top"):
            self.root.after(
                100,
                frame.scroll_to_top,
            )

    # ======================================================
    # BIÊN BẢN NGHIỆM THU
    # ======================================================

    def show_nghiem_thu(self):
        self.clear_content()
        frame = NghiemThuFrame(self.content)
        frame.pack(
            fill="both",
            expand=True,
        )

    # ======================================================
    # SETTINGS
    # ======================================================

    def show_settings(self):
        self.clear_content()
        ttk.Label(
            self.content,
            text="Cài đặt",
            font=("Arial", 20, "bold"),
        ).pack(pady=20)
        ttk.Label(
            self.content,
            text="Khu vực cài đặt hệ thống",
        ).pack()

    # ======================================================
    # OPEN DIRECTORY
    # ======================================================

    def open_dir(self, path):
        directory = os.path.abspath(path)
        if not os.path.exists(directory):
            os.makedirs(directory)
        try:
            os.startfile(directory)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(
                "Lỗi",
                f"Không thể mở thư mục:\n{directory}\n\n{exc}",
            )

    # ======================================================
    # CLEAR OUTPUTS
    # ======================================================

    def clear_outputs(self):
        output_dir = os.path.abspath("./outputs")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

            messagebox.showinfo(
                "Thông báo",
                "Thư mục outputs chưa có file nào.",
            )

            return

        files = []

        for filename in os.listdir(output_dir):
            path = os.path.join(
                output_dir,
                filename,
            )

            if os.path.isfile(path):
                files.append(path)

        if not files:
            messagebox.showinfo(
                "Thông báo",
                "Thư mục outputs không có file nào.",
            )

            return

        result = messagebox.askyesno(
            "Xác nhận xóa",
            f"Bạn có chắc muốn xóa {len(files)} file trong thư mục outputs?",
        )

        if not result:
            return

        deleted = 0
        errors = []

        for path in files:
            try:
                os.remove(path)
                deleted += 1

            except Exception:  # noqa: BLE001
                errors.append(
                    f"{os.path.basename(path)}: File đang được sử dụng, không thể xoá!"
                )

        if errors:
            messagebox.showwarning(
                "Hoàn tất",
                f"Đã xóa {deleted}/{len(files)} file.\n\n"
                "Không thể xóa:\n" + "\n".join(errors),
            )

        else:
            messagebox.showinfo(
                "Hoàn tất",
                f"Đã xóa thành công {deleted} file.",
            )

    # ======================================================
    # EXIT
    # ======================================================

    def exit_app(self):
        result = messagebox.askyesno(
            "Thoát",
            "Bạn có muốn thoát chương trình?",
        )

        if result:
            self.root.destroy()


# ==========================================================
# MAIN
# ==========================================================

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = MainApp(root)
    root.mainloop()
