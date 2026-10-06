import os
import tkinter as tk
from tkinter import messagebox, ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

from hop_dong_duyet_gia import HopDongDuyetGiaFrame
from NghiemThu_scrollbar_fixed import NghiemThuFrame


class MainApp:
    def __init__(self, root):
        self.root = root

        self.root.title("Hệ thống quản lý tài liệu")
        self.root.geometry("1200x750")
        self.root.minsize(1200, 750)

        # ==================================================
        # STYLE
        # ==================================================

        # ==================================================
        # STYLE DÙNG CHUNG TOÀN BỘ ỨNG DỤNG
        # ==================================================

        style = ttk.Style()

        try:
            style.theme_use("vista")
        except Exception:  # noqa: BLE001, S110
            pass

        # ------------------------------
        # NỀN CHUNG
        # ------------------------------

        style.configure(
            "TFrame",
            background="#f2f2f2",
        )

        style.configure(
            "TLabelframe",
            background="#f2f2f2",
        )

        style.configure(
            "TLabelframe.Label",
            background="#f2f2f2",
            font=("Arial", 10),
        )

        # ------------------------------
        # LABEL
        # ------------------------------

        style.configure(
            "TLabel",
            background="#f2f2f2",
            font=("Arial", 10),
        )

        style.configure(
            "AppTitle.TLabel",
            background="#f2f2f2",
            font=("Arial", 22, "bold"),
        )

        style.configure(
            "PageTitle.TLabel",
            background="#f2f2f2",
            font=("Arial", 18, "bold"),
        )

        style.configure(
            "Section.TLabel",
            background="#f2f2f2",
            font=("Arial", 11, "bold"),
        )

        style.configure(
            "Header.TLabel",
            background="#f2f2f2",
            font=("Arial", 10, "bold"),
        )

        # ------------------------------
        # BUTTON
        # ------------------------------

        style.configure(
            "TButton",
            font=("Arial", 10),
            padding=(8, 4),
        )

        style.configure(
            "Menu.TButton",
            font=("Arial", 10),
            padding=(10, 8),
        )

        style.configure(
            "Action.TButton",
            font=("Arial", 10, "bold"),
            padding=(10, 6),
        )

        # ------------------------------
        # RADIOBUTTON
        # ------------------------------

        style.configure(
            "TRadiobutton",
            background="#f2f2f2",
            font=("Arial", 10),
        )

        # ------------------------------
        # ENTRY
        # ------------------------------

        style.configure(
            "TEntry",
            padding=(4, 3),
        )

        # ------------------------------
        # TREEVIEW
        # ------------------------------

        style.configure(
            "Treeview",
            font=("Arial", 9),
            rowheight=24,
        )

        style.configure(
            "Treeview.Heading",
            font=("Arial", 9, "bold"),
            padding=(4, 4),
        )
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

        sidebar = ttk.LabelFrame(
            body,
            text="Chức năng",
        )

        sidebar.pack(
            side="left",
            fill="y",
            padx=(0, 15),
        )

        # ==================================================
        # CONTENT
        # ==================================================
        #
        # QUAN TRỌNG:
        # Không tạo Canvas/Scrollbar ở đây.
        #
        # Các màn hình con như HopDongDuyetGiaFrame
        # sẽ tự quản lý scrollbar của chúng.
        #
        # Tránh lỗi scrollbar lồng nhau.
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
        # MENU
        # ==================================================

        ttk.Button(
            sidebar,
            text="📄  Tạo duyệt giá - hợp đồng",
            style="Menu.TButton",
            command=self.show_documents,
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="📋  Biên bản nghiệm thu",
            style="Menu.TButton",
            command=self.show_nghiem_thu,
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="⚙  Cài đặt",
            style="Menu.TButton",
            command=self.show_settings,
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="📁 Mở thư mục templates",
            style="Menu.TButton",
            command=lambda: self.open_dir("./templates"),
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="📁 Mở thư mục outputs",
            style="Menu.TButton",
            command=lambda: self.open_dir("./outputs"),
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="🗑  Xóa file outputs",
            style="Menu.TButton",
            command=self.clear_outputs,
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        ttk.Button(
            sidebar,
            text="❌  Thoát",
            style="Menu.TButton",
            command=self.exit_app,
        ).pack(
            fill="x",
            padx=10,
            pady=10,
        )

        # ==================================================
        # HOME
        # ==================================================

        self.show_home()

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
