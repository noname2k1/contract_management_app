# nghiem_thu_frame:
import os
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from utils import (
    calculate_service_totals,
    flatten_jobs,
    format_money,
    generate_docx,
    number_to_vietnamese_words,
    open_file,
    parse_number,
)

# ============================================================
# SCROLLABLE FRAME
# ============================================================
# Không còn canvas lồng với main.py.
# Chỉ màn hình này quản lý scrollbar.
# ============================================================


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


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_SERVICE_TEMPLATE = os.path.join(
    BASE_DIR, "templates", "bb_nghiem_thu_hd_dich_vu.docx"
)
DEFAULT_THANH_LI_TEMPLATE = os.path.join(
    BASE_DIR, "templates", "bb_thanh_ly_hd_dich_vu.docx"
)
DEFAULT_SALE_TEMPLATE = os.path.join(
    BASE_DIR, "templates", "bb_nghiem_thu_hd_mua_ban.docx"
)
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")


def safe_filename(value):
    """Chuẩn hóa chuỗi dùng làm tên file."""
    value = str(value or "").strip()
    for char in '<>:"/\\|?*':
        value = value.replace(char, "-")
    return value or "khong_so"


def parse_date(value, field_name):
    value = str(value or "").strip()
    try:
        return datetime.strptime(value, "%d/%m/%Y")  # noqa: DTZ007
    except ValueError as exc:
        raise ValueError(f"{field_name} phải có dạng DD/MM/YYYY.") from exc


# ============================================================
# MAIN FRAME
# ============================================================


class NghiemThuFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.loai_hd = tk.StringVar(value="Dịch vụ")

        self.jobs = []
        self.my_group = []
        self.partner_group = []
        self.products = []

        # Không reset template đã chọn khi đổi loại hợp đồng.
        self.template_file = DEFAULT_SERVICE_TEMPLATE
        self.template_file_thanh_li = DEFAULT_THANH_LI_TEMPLATE
        self.sale_template_file = DEFAULT_SALE_TEMPLATE

        # ----------------------------------------------------
        # COMMON
        # ----------------------------------------------------
        self.ma_hd = tk.StringVar(value="08220926/HĐKT/CK83/ITG/2026")
        self.ngay_hd = tk.StringVar(value="22/09/2026")
        self.ngay_nt = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))  # noqa: DTZ005
        self.ten_doi_tac = tk.StringVar(value="Công Ty Cổ Phần Công Nghệ ITG")

        # ----------------------------------------------------
        # SERVICE
        # ----------------------------------------------------
        self.ma_nt = tk.StringVar(value="NT-08220926/2026")
        self.ma_tl = tk.StringVar(value="TL-08220926/2026")
        self.dia_chi_doi_tac = tk.StringVar(
            value="Tầng 14, Tòa nhà Ladeco, 266 Đội Cấn, Ba Đình, Hà Nội"
        )
        self.sdt_doi_tac = tk.StringVar(value="024 3833 9966")
        self.ma_so_thue_doi_tac = tk.StringVar(value="0102345678")
        self.dai_dien_doi_tac = tk.StringVar(value="Nguyễn Văn A")
        self.chuc_vu_dai_dien_doi_tac = tk.StringVar(value="Tổng Giám đốc")
        self.stk_doi_tac = tk.StringVar(value="1220989888")
        self.ngan_hang_doi_tac = tk.StringVar(value="Ngân hàng BIDV")
        self.chi_nhanh_bank_doi_tac = tk.StringVar(value="Chi nhánh Hà Thành")
        self.tien_da_thanh_toan = tk.StringVar(value="0")

        # ----------------------------------------------------
        # SALE
        # ----------------------------------------------------
        self.ten_bbnt = tk.StringVar(value="NGHIỆM THU HỢP ĐỒNG MUA BÁN")
        self.hang_muc_nt_hdmb = tk.StringVar(value="")
        self.gio_bat_dau_nt = tk.StringVar(value="08:00")
        self.gio_ket_thuc_nt = tk.StringVar(value="17:00")
        self.muc_dich_hdmb = tk.StringVar(value="")

        self.build_ui()
        self.load_sample_jobs()
        self.load_sample_sale_data()
        self.change_contract_type()

    # ========================================================
    # ERROR HANDLING
    # ========================================================

    def report_callback_exception(self, exc, val, tb):
        error_text = "".join(traceback.format_exception(exc, val, tb))
        print("\n" + "=" * 80)
        print("LỖI CHƯƠNG TRÌNH")
        print("=" * 80)
        print(error_text)
        print("=" * 80)

        try:
            self.clipboard_clear()
            self.clipboard_append(error_text)
            self.update()
        except Exception:
            pass

        messagebox.showerror(
            "Lỗi chương trình",
            "Đã xảy ra lỗi trong chương trình.\n\n"
            "Chi tiết lỗi đã được copy vào Clipboard.",
            parent=self,
        )

    def build_ui(self):
        # ========================================================
        # CHỈ DÙNG MỘT SCROLLABLE FRAME
        # Giống HopDongDuyetGiaFrame:
        # NghiemThuFrame tự quản lý canvas + scrollbar.
        # ========================================================

        self.main = ScrollableFrame(self)
        self.main.pack(
            fill="both",
            expand=True,
            padx=0,
            pady=0,
        )

        # Toàn bộ nội dung của màn hình nằm trong inner.
        self.form = self.main.inner

        title_frame = ttk.Frame(self.form)
        title_frame.pack(fill="x", padx=10, pady=(5, 2))

        ttk.Label(
            title_frame,
            text="QUẢN LÝ BIÊN BẢN NGHIỆM THU",
            font=("Arial", 16, "bold"),
        ).pack(side="left")

        # Chọn file / xuất file
        self.bottom = ttk.Frame(self.form)
        self.bottom.pack(fill="x", padx=10, pady=5)

        self.template_row = ttk.Frame(self.bottom)
        self.template_row.pack(fill="x", pady=2)

        self.template_label = ttk.Label(
            self.template_row,
            text=self.template_file,
            anchor="w",
        )
        self.template_label.pack(side="left", fill="x", expand=True)

        ttk.Button(
            self.template_row,
            text="📁 Template nghiệm thu",
            style="NTFile.TButton",
            command=self.select_template,
        ).pack(side="left", padx=5)

        self.template_thanh_li_frame = ttk.Frame(self.template_row)
        self.template_thanh_li_frame.pack(
            side="left",
            fill="x",
            expand=True,
        )

        self.template_thanh_li_label = ttk.Label(
            self.template_thanh_li_frame,
            text=self.template_file_thanh_li,
            anchor="w",
        )
        self.template_thanh_li_label.pack(
            side="left",
            fill="x",
            expand=True,
        )

        ttk.Button(
            self.template_thanh_li_frame,
            text="📁 Template thanh lý",
            style="NTFile.TButton",
            command=self.select_template_thanh_li,
        ).pack(side="right", padx=5)

        self.export_row = ttk.Frame(self.bottom)
        self.export_row.pack(fill="x", pady=5)

        ttk.Button(
            self.export_row,
            text="📄 Xuất Word",
            style="NTSuccess.TButton",
            command=self.export_docx,
        ).pack(side="right", padx=5)

        # Loại hợp đồng
        self.info_frame = ttk.LabelFrame(
            self.form,
            text="Thông tin chung",
        )
        self.info_frame.pack(fill="x", padx=10, pady=5)

        type_frame = ttk.LabelFrame(
            self.form,
            text="Loại hợp đồng",
        )
        type_frame.pack(fill="x", padx=10, pady=5)

        ttk.Radiobutton(
            type_frame,
            text="Dịch vụ",
            variable=self.loai_hd,
            value="Dịch vụ",
            command=self.change_contract_type,
        ).pack(side="left", padx=20, pady=5)

        ttk.Radiobutton(
            type_frame,
            text="Mua bán",
            variable=self.loai_hd,
            value="Mua bán",
            command=self.change_contract_type,
        ).pack(side="left", padx=20, pady=5)

        self.build_common_info()

        self.service_frame = ttk.Frame(self.form)
        self.build_service_form()

        self.sale_frame = ttk.Frame(self.form)
        self.build_sale_form()

        # Luôn bắt đầu ở đầu form sau khi tạo xong nội dung.
        self.after_idle(self.scroll_to_top)

    def scroll_to_top(self):
        if hasattr(self, "main") and hasattr(self.main, "canvas"):
            self.main.canvas.yview_moveto(0)

    # ========================================================
    # COMMON INFO
    # ========================================================

    def build_common_info(self):
        frame = self.info_frame

        for col in range(4):
            frame.columnconfigure(col, weight=1)

        self.create_label(frame, "Số hợp đồng:", 0, 0)
        ttk.Entry(frame, textvariable=self.ma_hd).grid(
            row=0, column=1, sticky="ew", padx=5, pady=3
        )

        self.create_label(frame, "Ngày hợp đồng:", 0, 2)
        ttk.Entry(frame, textvariable=self.ngay_hd).grid(
            row=0, column=3, sticky="ew", padx=5, pady=3
        )

        self.create_label(frame, "Ngày nghiệm thu:", 1, 0)
        ttk.Entry(frame, textvariable=self.ngay_nt).grid(
            row=1, column=1, sticky="ew", padx=5, pady=3
        )

        self.create_label(frame, "Tên đối tác:", 1, 2)
        ttk.Entry(frame, textvariable=self.ten_doi_tac).grid(
            row=1, column=3, sticky="ew", padx=5, pady=3
        )

        self.service_info_frame = ttk.Frame(frame)
        self.service_info_frame.grid(row=2, column=0, columnspan=4, sticky="ew")

        for col in range(4):
            self.service_info_frame.columnconfigure(col, weight=1)

        self.create_label(self.service_info_frame, "Số biên bản:", 0, 0)
        ttk.Entry(self.service_info_frame, textvariable=self.ma_nt).grid(
            row=0, column=1, sticky="ew", padx=5, pady=3
        )

        self.create_label(self.service_info_frame, "Mã thanh lý:", 0, 2)
        ttk.Entry(self.service_info_frame, textvariable=self.ma_tl).grid(
            row=0, column=3, sticky="ew", padx=5, pady=3
        )

        self.create_label(self.service_info_frame, "Địa chỉ đối tác:", 1, 0)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.dia_chi_doi_tac,
        ).grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Số điện thoại:", 2, 0)
        ttk.Entry(self.service_info_frame, textvariable=self.sdt_doi_tac).grid(
            row=2, column=1, sticky="ew", padx=5, pady=3
        )

        self.create_label(self.service_info_frame, "Mã số thuế:", 2, 2)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.ma_so_thue_doi_tac,
        ).grid(row=2, column=3, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Đại diện:", 3, 0)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.dai_dien_doi_tac,
        ).grid(row=3, column=1, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Chức vụ:", 3, 2)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.chuc_vu_dai_dien_doi_tac,
        ).grid(row=3, column=3, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Số tài khoản:", 4, 0)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.stk_doi_tac,
        ).grid(row=4, column=1, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Ngân hàng:", 4, 2)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.ngan_hang_doi_tac,
        ).grid(row=4, column=3, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Chi nhánh:", 5, 0)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.chi_nhanh_bank_doi_tac,
        ).grid(row=5, column=1, columnspan=3, sticky="ew", padx=5, pady=3)

        self.create_label(self.service_info_frame, "Đã thanh toán:", 6, 0)
        ttk.Entry(
            self.service_info_frame,
            textvariable=self.tien_da_thanh_toan,
        ).grid(row=6, column=1, sticky="ew", padx=5, pady=3)

    # ========================================================
    # SERVICE FORM
    # ========================================================

    def build_service_form(self):
        job_frame = ttk.LabelFrame(self.service_frame, text="Danh sách công việc")
        job_frame.pack(fill="both", expand=True, padx=10, pady=5)

        toolbar = ttk.Frame(job_frame)
        toolbar.pack(fill="x", pady=3)

        ttk.Button(
            toolbar,
            text="➕ Thêm công việc",
            style="NTSuccess.TButton",
            command=self.add_job,
        ).pack(side="left", padx=3)

        ttk.Button(
            toolbar,
            text="➕ Thêm công việc con",
            style="NTPrimary.TButton",
            command=self.add_sub_job,
        ).pack(side="left", padx=3)

        ttk.Button(
            toolbar, text="✏ Sửa", style="NTWarning.TButton", command=self.edit_job
        ).pack(side="left", padx=3)

        ttk.Button(
            toolbar, text="🗑 Xóa", style="NTDanger.TButton", command=self.delete_job
        ).pack(side="left", padx=3)

        tree_frame = ttk.Frame(job_frame)
        tree_frame.pack(fill="both", expand=True)

        columns = (
            "stt",
            "ten_dich_vu",
            "cong_viec_thuc_hien",
            "don_vi",
            "so_luong",
            "don_gia",
            "thanh_tien",
        )

        self.job_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="tree headings",
            height=10,
        )

        self.job_tree.heading("#0", text="")
        self.job_tree.column("#0", width=22, minwidth=22, stretch=False)

        headings = {
            "stt": "STT",
            "ten_dich_vu": "Tên dịch vụ",
            "cong_viec_thuc_hien": "Công việc thực hiện",
            "don_vi": "Đơn vị",
            "so_luong": "Số lượng",
            "don_gia": "Đơn giá",
            "thanh_tien": "Thành tiền",
        }

        widths = {
            "stt": 45,
            "ten_dich_vu": 220,
            "cong_viec_thuc_hien": 260,
            "don_vi": 70,
            "so_luong": 75,
            "don_gia": 110,
            "thanh_tien": 120,
        }

        for col in columns:
            self.job_tree.heading(col, text=headings[col])
            self.job_tree.column(
                col,
                width=widths[col],
                anchor="center"
                if col in {"stt", "don_vi", "so_luong"}
                else "e"
                if col in {"don_gia", "thanh_tien"}
                else "w",
            )

        scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.job_tree.yview,
        )
        self.job_tree.configure(yscrollcommand=scrollbar.set)

        # QUAN TRỌNG: tree và scrollbar cùng dùng tree_frame.
        self.job_tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y", pady=5)

        self.job_tree.bind("<Double-1>", lambda _event: self.edit_job())

    # ========================================================
    # SALE FORM
    # ========================================================

    def build_sale_form(self):
        basic_frame = ttk.LabelFrame(
            self.sale_frame,
            text="Thông tin biên bản mua bán",
        )
        basic_frame.pack(fill="x", padx=10, pady=5)

        for col in range(4):
            basic_frame.columnconfigure(col, weight=1)

        self.create_label(basic_frame, "Tên biên bản:", 0, 0)
        ttk.Entry(basic_frame, textvariable=self.ten_bbnt).grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=3
        )

        self.create_label(basic_frame, "Hạng mục nghiệm thu:", 1, 0)
        ttk.Entry(basic_frame, textvariable=self.hang_muc_nt_hdmb).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=3
        )

        self.create_label(basic_frame, "Mục đích hợp đồng:", 2, 0)
        ttk.Entry(basic_frame, textvariable=self.muc_dich_hdmb).grid(
            row=2, column=1, columnspan=3, sticky="ew", padx=5, pady=3
        )

        self.create_label(basic_frame, "Bắt đầu:", 3, 0)
        ttk.Entry(basic_frame, textvariable=self.gio_bat_dau_nt, width=15).grid(
            row=3, column=1, sticky="w", padx=5, pady=3
        )

        self.create_label(basic_frame, "Kết thúc:", 3, 2)
        ttk.Entry(basic_frame, textvariable=self.gio_ket_thuc_nt, width=15).grid(
            row=3, column=3, sticky="w", padx=5, pady=3
        )

        member_container = ttk.Frame(self.sale_frame)
        member_container.pack(fill="both", expand=True, padx=10, pady=5)

        member_container.columnconfigure(0, weight=1)
        member_container.columnconfigure(1, weight=1)
        member_container.rowconfigure(0, weight=1)

        my_frame = ttk.LabelFrame(
            member_container,
            text="Thành phần Công ty TNHH MTV Cơ khí 83",
        )
        my_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.build_member_tree(my_frame, "my")

        partner_frame = ttk.LabelFrame(member_container, text="Thành phần đối tác")
        partner_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.build_member_tree(partner_frame, "partner")

        product_frame = ttk.LabelFrame(
            self.sale_frame,
            text="Danh sách hàng hóa / thiết bị",
        )
        product_frame.pack(fill="both", expand=True, padx=10, pady=5)

        toolbar = ttk.Frame(product_frame)
        toolbar.pack(fill="x", pady=3)

        ttk.Button(
            toolbar,
            text="➕ Thêm sản phẩm",
            style="NTSuccess.TButton",
            command=self.add_product,
        ).pack(side="left", padx=3)

        ttk.Button(
            toolbar, text="✏ Sửa", style="NTWarning.TButton", command=self.edit_product
        ).pack(side="left", padx=3)

        ttk.Button(
            toolbar, text="🗑 Xóa", style="NTDanger.TButton", command=self.delete_product
        ).pack(side="left", padx=3)

        product_tree_frame = ttk.Frame(product_frame)
        product_tree_frame.pack(fill="both", expand=True)

        columns = ("stt", "name", "unit", "sl_hd", "sl_thucte", "chenh_lech")

        self.product_tree = ttk.Treeview(
            product_tree_frame,
            columns=columns,
            show="headings",
            height=7,
        )

        headings = {
            "stt": "STT",
            "name": "Nội dung",
            "unit": "ĐVT",
            "sl_hd": "Hợp đồng",
            "sl_thucte": "Thực tế",
            "chenh_lech": "Chênh lệch",
        }

        for col, heading in headings.items():
            self.product_tree.heading(col, text=heading)

        self.product_tree.column(
            "stt", width=45, minwidth=40, anchor="center", stretch=False
        )

        self.product_tree.column(
            "name", width=300, minwidth=180, anchor="w", stretch=True
        )

        self.product_tree.column(
            "unit", width=70, minwidth=55, anchor="center", stretch=False
        )

        self.product_tree.column(
            "sl_hd", width=90, minwidth=70, anchor="center", stretch=False
        )

        self.product_tree.column(
            "sl_thucte", width=90, minwidth=70, anchor="center", stretch=False
        )

        self.product_tree.column(
            "chenh_lech", width=90, minwidth=70, anchor="center", stretch=False
        )

        product_scroll = ttk.Scrollbar(
            product_tree_frame,
            orient="vertical",
            command=self.product_tree.yview,
        )
        self.product_tree.configure(yscrollcommand=product_scroll.set)

        self.product_tree.pack(side="left", fill="both", expand=True)
        product_scroll.pack(side="right", fill="y")

        self.product_tree.bind("<Double-1>", lambda _event: self.edit_product())

    # ========================================================
    # MEMBER TREE
    # ========================================================

    def build_member_tree(self, parent, group_type):
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=3)

        if group_type == "my":
            ttk.Button(
                toolbar,
                text="➕ Thêm",
                style="NTSuccess.TButton",
                command=self.add_my_member,
            ).pack(side="left", padx=2)
            ttk.Button(
                toolbar,
                text="✏ Sửa",
                style="NTWarning.TButton",
                command=self.edit_my_member,
            ).pack(side="left", padx=2)
            ttk.Button(
                toolbar,
                text="🗑 Xóa",
                style="NTDanger.TButton",
                command=self.delete_my_member,
            ).pack(side="left", padx=2)
        else:
            ttk.Button(
                toolbar,
                text="➕ Thêm",
                style="NTSuccess.TButton",
                command=self.add_partner_member,
            ).pack(side="left", padx=2)
            ttk.Button(
                toolbar,
                text="✏ Sửa",
                style="NTWarning.TButton",
                command=self.edit_partner_member,
            ).pack(side="left", padx=2)
            ttk.Button(
                toolbar,
                text="🗑 Xóa",
                style="NTDanger.TButton",
                command=self.delete_partner_member,
            ).pack(side="left", padx=2)

        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(
            tree_frame,
            columns=("stt", "name", "role"),
            show="headings",
            height=6,
        )

        tree.heading("stt", text="STT")
        tree.heading("name", text="Họ tên")
        tree.heading("role", text="Chức vụ")

        tree.column("stt", width=45, minwidth=40, anchor="center", stretch=False)
        tree.column("name", width=180, minwidth=120, anchor="w", stretch=True)
        tree.column("role", width=180, minwidth=120, anchor="w", stretch=True)

        scroll = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=scroll.set)

        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        tree.bind(
            "<Double-1>",
            lambda _event, t=group_type: self.edit_member_by_type(t),
        )

        if group_type == "my":
            self.my_tree = tree
        else:
            self.partner_tree = tree

    # ========================================================
    # CONTRACT TYPE
    # ========================================================

    def change_contract_type(self):
        if self.loai_hd.get() == "Dịch vụ":
            self.sale_frame.pack_forget()

            self.service_frame.pack(fill="both", expand=True, padx=0, pady=0)

            self.service_info_frame.grid()
            self.template_thanh_li_frame.pack(side="left", fill="x", expand=True)

            self.template_file = self.template_file or (DEFAULT_SERVICE_TEMPLATE)
            self.template_label.config(text=self.template_file)
        else:
            self.service_frame.pack_forget()

            self.sale_frame.pack(fill="both", expand=True, padx=0, pady=0)

            self.service_info_frame.grid_remove()
            self.template_thanh_li_frame.pack_forget()

            self.template_file = self.sale_template_file or DEFAULT_SALE_TEMPLATE
            self.template_label.config(text=self.template_file)

    # ========================================================
    # HELPERS
    # ========================================================

    def create_label(self, parent, text, row, column):
        ttk.Label(parent, text=text).grid(
            row=row,
            column=column,
            sticky="w",
            padx=5,
            pady=3,
        )

    # ========================================================
    # SAMPLE DATA
    # ========================================================

    def load_sample_jobs(self):
        self.jobs = [
            {
                "stt": "1",
                "ten_dich_vu": "Khảo sát và tư vấn thiết bị IoT",
                "cong_viec_thuc_hien": "",
                "don_vi": "Gói",
                "so_luong": 1,
                "don_gia": 15000000,
                "thanh_tien": 15000000,
                "children": [
                    {
                        "stt": "1.1",
                        "ten_dich_vu": "Khảo sát",
                        "cong_viec_thuc_hien": "Khảo sát hiện trạng hệ thống.",
                    },
                    {
                        "stt": "1.2",
                        "ten_dich_vu": "Tư vấn",
                        "cong_viec_thuc_hien": "Tư vấn thiết bị phù hợp.",
                    },
                ],
            },
            {
                "stt": "2",
                "ten_dich_vu": "Cấu hình và kết nối tín hiệu hệ thống với mạng TSLqs",
                "cong_viec_thuc_hien": "",
                "don_vi": "Gói",
                "so_luong": 1,
                "don_gia": 25000000,
                "thanh_tien": 25000000,
                "children": [],
            },
            {
                "stt": "3",
                "ten_dich_vu": "Đào tạo và bàn giao vận hành",
                "cong_viec_thuc_hien": "",
                "don_vi": "Gói",
                "so_luong": 1,
                "don_gia": 15000000,
                "thanh_tien": 15000000,
                "children": [
                    {
                        "stt": "3.1",
                        "ten_dich_vu": "Đào tạo",
                        "cong_viec_thuc_hien": "Đào tạo vận hành.",
                    },
                    {
                        "stt": "3.2",
                        "ten_dich_vu": "Bàn giao",
                        "cong_viec_thuc_hien": "Bàn giao hệ thống.",
                    },
                ],
            },
        ]
        self.refresh_job_tree()

    def load_sample_sale_data(self):
        self.ma_hd.set("08220926/HĐMB/CK83/ITG/2026")
        self.ngay_hd.set("22/09/2026")
        self.ngay_nt.set(datetime.now().strftime("%d/%m/%Y"))  # noqa: DTZ005
        self.ten_doi_tac.set("Công Ty Cổ Phần Công Nghệ ITG")

        self.ten_bbnt.set("NGHIỆM THU HỢP ĐỒNG MUA BÁN THIẾT BỊ")
        self.hang_muc_nt_hdmb.set("Cung cấp và lắp đặt thiết bị công nghệ thông tin")
        self.gio_bat_dau_nt.set("08:00")
        self.gio_ket_thuc_nt.set("17:00")
        self.muc_dich_hdmb.set(
            "Cung cấp thiết bị phục vụ triển khai hệ thống công nghệ thông tin"
        )

        self.my_group = [
            {"stt": 1, "name": "Vũ Việt Thắng", "role": "Phó Giám đốc"},
            {"stt": 2, "name": "Nguyễn Văn Minh", "role": "Trưởng phòng Kỹ thuật"},
            {"stt": 3, "name": "Trần Văn Hùng", "role": "Cán bộ kỹ thuật"},
        ]

        self.partner_group = [
            {"stt": 1, "name": "Nguyễn Văn A", "role": "Tổng Giám đốc"},
            {"stt": 2, "name": "Lê Văn B", "role": "Giám đốc Kỹ thuật"},
            {"stt": 3, "name": "Phạm Văn C", "role": "Kỹ sư triển khai"},
        ]

        self.products = [
            {
                "stt": 1,
                "name": "Máy tính để bàn",
                "unit": "Bộ",
                "sl_hd": 10,
                "sl_thucte": 10,
                "chenh_lech": 0,
            },
            {
                "stt": 2,
                "name": "Switch mạng lớp 2",
                "unit": "Cái",
                "sl_hd": 5,
                "sl_thucte": 5,
                "chenh_lech": 0,
            },
            {
                "stt": 3,
                "name": "Máy Scan tài liệu",
                "unit": "Cái",
                "sl_hd": 3,
                "sl_thucte": 3,
                "chenh_lech": 0,
            },
            {
                "stt": 4,
                "name": "Bộ lưu điện UPS",
                "unit": "Bộ",
                "sl_hd": 5,
                "sl_thucte": 4,
                "chenh_lech": -1,
            },
            {
                "stt": 5,
                "name": "Tủ mạng 42U",
                "unit": "Tủ",
                "sl_hd": 2,
                "sl_thucte": 2,
                "chenh_lech": 0,
            },
        ]

        self.refresh_my_group()
        self.refresh_partner_group()
        self.refresh_products()

    # ========================================================
    # JOBS
    # ========================================================

    def refresh_job_tree(self):
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)

        for job in self.jobs:
            parent_id = self.job_tree.insert(
                "",
                "end",
                values=(
                    job.get("stt", ""),
                    job.get("ten_dich_vu", ""),
                    job.get("cong_viec_thuc_hien", ""),
                    job.get("don_vi", "Gói"),
                    job.get("so_luong", 1),
                    format_money(job.get("don_gia", 0)),
                    format_money(job.get("thanh_tien", 0)),
                ),
                open=True,
            )

            for child in job.get("children", []):
                self.job_tree.insert(
                    parent_id,
                    "end",
                    values=(
                        child.get("stt", ""),
                        child.get("ten_dich_vu", ""),
                        child.get("cong_viec_thuc_hien", ""),
                        "",
                        "",
                        "",
                        "",
                    ),
                )

    def get_next_stt(self):
        return str(len(self.jobs) + 1)

    def add_job(self):
        dialog = JobDialog(self, "Thêm công việc")
        if not dialog.result:
            return

        data = dialog.result
        self.jobs.append(
            {
                "stt": self.get_next_stt(),
                **data,
                "children": [],
            }
        )
        self.refresh_job_tree()

    def add_sub_job(self):
        selection = self.job_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn công việc cha.",
                parent=self,
            )
            return

        item = selection[0]
        if self.job_tree.parent(item):
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn công việc cấp cha.",
                parent=self,
            )
            return

        parent_stt = self.job_tree.item(item, "values")[0]
        parent_job, _child = self.find_job_by_stt(parent_stt)

        if not parent_job:
            return

        dialog = JobDialog(self, "Thêm công việc con")
        if not dialog.result:
            return

        child_number = len(parent_job.get("children", [])) + 1
        child_stt = f"{parent_stt}.{child_number}"

        parent_job.setdefault("children", []).append(
            {
                "stt": child_stt,
                "ten_dich_vu": dialog.result["ten_dich_vu"],
                "cong_viec_thuc_hien": dialog.result["cong_viec_thuc_hien"],
            }
        )
        self.refresh_job_tree()

    def edit_job(self):
        selection = self.job_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn công việc.",
                parent=self,
            )
            return

        stt = self.job_tree.item(selection[0], "values")[0]
        job, child = self.find_job_by_stt(stt)

        if not job:
            return

        current = child if child else job
        dialog = JobDialog(self, "Sửa công việc", current)
        if not dialog.result:
            return

        current["ten_dich_vu"] = dialog.result["ten_dich_vu"]
        current["cong_viec_thuc_hien"] = dialog.result["cong_viec_thuc_hien"]

        if not child:
            current["don_vi"] = dialog.result["don_vi"]
            current["so_luong"] = dialog.result["so_luong"]
            current["don_gia"] = dialog.result["don_gia"]
            current["thanh_tien"] = dialog.result["thanh_tien"]

        self.refresh_job_tree()

    def find_job_by_stt(self, stt):
        for job in self.jobs:
            if str(job.get("stt")) == str(stt):
                return job, None

            for child in job.get("children", []):
                if str(child.get("stt")) == str(stt):
                    return job, child

        return None, None

    def delete_job(self):
        selection = self.job_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn công việc.",
                parent=self,
            )
            return

        stt = self.job_tree.item(selection[0], "values")[0]
        parent_job, child = self.find_job_by_stt(stt)

        if not parent_job:
            return

        if not messagebox.askyesno(
            "Xác nhận",
            f"Bạn có chắc muốn xóa công việc {stt}?",
            parent=self,
        ):
            return

        if child:
            parent_job["children"].remove(child)
        else:
            self.jobs.remove(parent_job)
            self.reindex_jobs()

        self.refresh_job_tree()

    def reindex_jobs(self):
        for index, job in enumerate(self.jobs, start=1):
            old_stt = str(job.get("stt", ""))
            job["stt"] = str(index)
            for child_index, child in enumerate(job.get("children", []), start=1):
                child["stt"] = f"{index}.{child_index}"

    # ========================================================
    # MEMBERS
    # ========================================================

    def add_my_member(self):
        dialog = MemberDialog(self, "Thêm thành viên")
        if not dialog.result:
            return

        self.my_group.append(
            {
                "stt": len(self.my_group) + 1,
                **dialog.result,
            }
        )
        self.refresh_my_group()

    def edit_my_member(self):
        selection = self.my_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn thành viên.",
                parent=self,
            )
            return

        index = self.my_tree.index(selection[0])
        member = self.my_group[index]

        dialog = MemberDialog(self, "Sửa thành viên", member)
        if not dialog.result:
            return

        member.update(dialog.result)
        self.refresh_my_group()

    def delete_my_member(self):
        selection = self.my_tree.selection()
        if not selection:
            return

        index = self.my_tree.index(selection[0])

        if not messagebox.askyesno(
            "Xác nhận",
            "Bạn có chắc muốn xóa thành viên này?",
            parent=self,
        ):
            return

        del self.my_group[index]
        self.reindex_members(self.my_group)
        self.refresh_my_group()

    def add_partner_member(self):
        dialog = MemberDialog(self, "Thêm thành viên đối tác")
        if not dialog.result:
            return

        self.partner_group.append(
            {
                "stt": len(self.partner_group) + 1,
                **dialog.result,
            }
        )
        self.refresh_partner_group()

    def edit_partner_member(self):
        selection = self.partner_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn thành viên.",
                parent=self,
            )
            return

        index = self.partner_tree.index(selection[0])
        member = self.partner_group[index]

        dialog = MemberDialog(self, "Sửa thành viên", member)
        if not dialog.result:
            return

        member.update(dialog.result)
        self.refresh_partner_group()

    def delete_partner_member(self):
        selection = self.partner_tree.selection()
        if not selection:
            return

        index = self.partner_tree.index(selection[0])

        if not messagebox.askyesno(
            "Xác nhận",
            "Bạn có chắc muốn xóa thành viên này?",
            parent=self,
        ):
            return

        del self.partner_group[index]
        self.reindex_members(self.partner_group)
        self.refresh_partner_group()

    def edit_member_by_type(self, group_type):
        if group_type == "my":
            self.edit_my_member()
        else:
            self.edit_partner_member()

    @staticmethod
    def reindex_members(members):
        for index, member in enumerate(members, start=1):
            member["stt"] = index

    def refresh_my_group(self):
        for item in self.my_tree.get_children():
            self.my_tree.delete(item)

        for member in self.my_group:
            self.my_tree.insert(
                "",
                "end",
                values=(
                    member["stt"],
                    member["name"],
                    member["role"],
                ),
            )

    def refresh_partner_group(self):
        for item in self.partner_tree.get_children():
            self.partner_tree.delete(item)

        for member in self.partner_group:
            self.partner_tree.insert(
                "",
                "end",
                values=(
                    member["stt"],
                    member["name"],
                    member["role"],
                ),
            )

    # ========================================================
    # PRODUCTS
    # ========================================================

    @staticmethod
    def calculate_difference(sl_hd, sl_thucte):
        return parse_number(sl_thucte) - parse_number(sl_hd)

    def add_product(self):
        dialog = ProductDialog(self, "Thêm sản phẩm")
        if not dialog.result:
            return

        data = dialog.result
        self.products.append(
            {
                "stt": len(self.products) + 1,
                **data,
                "chenh_lech": self.calculate_difference(
                    data["sl_hd"],
                    data["sl_thucte"],
                ),
            }
        )
        self.refresh_products()

    def edit_product(self):
        selection = self.product_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng chọn sản phẩm.",
                parent=self,
            )
            return

        index = self.product_tree.index(selection[0])
        product = self.products[index]

        dialog = ProductDialog(self, "Sửa sản phẩm", product)
        if not dialog.result:
            return

        data = dialog.result
        product.update(data)
        product["chenh_lech"] = self.calculate_difference(
            data["sl_hd"],
            data["sl_thucte"],
        )
        self.refresh_products()

    def delete_product(self):
        selection = self.product_tree.selection()
        if not selection:
            return

        index = self.product_tree.index(selection[0])

        if not messagebox.askyesno(
            "Xác nhận",
            "Bạn có chắc muốn xóa sản phẩm này?",
            parent=self,
        ):
            return

        del self.products[index]
        for i, product in enumerate(self.products, start=1):
            product["stt"] = i

        self.refresh_products()

    def refresh_products(self):
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)

        for product in self.products:
            sl_hd = product.get("sl_hd", 0)
            sl_thucte = product.get("sl_thucte", 0)

            try:
                difference = self.calculate_difference(sl_hd, sl_thucte)
            except (TypeError, ValueError):
                difference = product.get("chenh_lech", 0)

            product["chenh_lech"] = difference

            self.product_tree.insert(
                "",
                "end",
                values=(
                    product["stt"],
                    product["name"],
                    product["unit"],
                    sl_hd,
                    sl_thucte,
                    difference,
                ),
            )

    # ========================================================
    # TEMPLATE
    # ========================================================

    def select_template(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file Word template",
            filetypes=[
                ("Word files", "*.docx"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return

        if self.loai_hd.get() == "Dịch vụ":
            self.template_file = file_path
        else:
            self.sale_template_file = file_path
            self.template_file = file_path

        self.template_label.config(text=file_path)

    def select_template_thanh_li(self):
        file_path = filedialog.askopenfilename(
            title="Chọn template thanh lý",
            filetypes=[
                ("Word Document", "*.docx"),
                ("All Files", "*.*"),
            ],
        )
        if not file_path:
            return

        self.template_file_thanh_li = file_path
        self.template_thanh_li_label.config(text=file_path)

    # ========================================================
    # VALIDATION
    # ========================================================

    def validate_common(self):
        if not self.ma_hd.get().strip():
            raise ValueError("Vui lòng nhập Số hợp đồng.")

        if not self.ngay_hd.get().strip():
            raise ValueError("Vui lòng nhập Ngày hợp đồng.")

        if not self.ngay_nt.get().strip():
            raise ValueError("Vui lòng nhập Ngày nghiệm thu.")

        parse_date(
            self.ngay_hd.get(),
            "Ngày hợp đồng",
        )
        parse_date(
            self.ngay_nt.get(),
            "Ngày nghiệm thu",
        )

    # ========================================================
    # EXPORT
    # ========================================================

    def export_docx(self):
        try:
            if self.loai_hd.get() == "Dịch vụ":
                self.export_service_docx()
            else:
                self.export_sale_docx()
        except Exception:  # noqa: BLE001
            self.report_callback_exception(*__import__("sys").exc_info())

    # ========================================================
    # EXPORT SERVICE
    # ========================================================

    def export_service_docx(self):
        self.validate_common()

        if not self.jobs:
            raise ValueError("Danh sách công việc đang trống.")

        if not os.path.exists(self.template_file):
            raise FileNotFoundError(
                "Không tìm thấy template nghiệm thu:\n"
                + os.path.abspath(self.template_file)
            )

        if not os.path.exists(self.template_file_thanh_li):
            raise FileNotFoundError(
                "Không tìm thấy template thanh lý:\n"
                + os.path.abspath(self.template_file_thanh_li)
            )

        ngay_nt = self.ngay_nt.get().strip()
        ngay_hd = self.ngay_hd.get().strip()

        date_obj_nt = parse_date(ngay_nt, "Ngày nghiệm thu")
        date_obj_hd = parse_date(ngay_hd, "Ngày hợp đồng")

        bang_cv_phang = flatten_jobs(self.jobs)

        # Bổ sung các trường display để template có thể dùng.
        for item in bang_cv_phang:
            item["don_gia_display"] = format_money(item.get("don_gia", 0))
            item["thanh_tien_display"] = format_money(item.get("thanh_tien", 0))

        tien_da_thanh_toan = parse_number(self.tien_da_thanh_toan.get())

        totals = calculate_service_totals(
            self.jobs,
            vat_rate=8,
            tien_da_thanh_toan=tien_da_thanh_toan,
        )

        tong_tien_chua_vat = totals["tong_tien_chua_vat"]
        vat = totals["VAT"]
        tong_tien_co_vat = totals["tong_tien_co_vat"]
        tien_con_phai_thanh_toan = totals["tien_con_phai_thanh_toan"]

        context = {
            "loai_hd": self.loai_hd.get(),
            "ma_nt": self.ma_nt.get().strip(),
            "ma_tl": self.ma_tl.get().strip(),
            "ma_hd": self.ma_hd.get().strip(),
            "ngay_hd": ngay_hd,
            "ngay_nt": ngay_nt,
            "date_nt": date_obj_nt.day,
            "month_nt": date_obj_nt.month,
            "year_nt": date_obj_nt.year,
            "date_hd": date_obj_hd.day,
            "month_hd": date_obj_hd.month,
            "year_hd": date_obj_hd.year,
            "ten_doi_tac": self.ten_doi_tac.get().strip(),
            "dia_chi_doi_tac": self.dia_chi_doi_tac.get().strip(),
            "sdt_doi_tac": self.sdt_doi_tac.get().strip(),
            "ma_so_thue_doi_tac": self.ma_so_thue_doi_tac.get().strip(),
            "dai_dien_doi_tac": self.dai_dien_doi_tac.get().strip(),
            "chuc_vu_dai_dien_doi_tac": self.chuc_vu_dai_dien_doi_tac.get().strip(),
            "stk_doi_tac": self.stk_doi_tac.get().strip(),
            "ngan_hang_doi_tac": self.ngan_hang_doi_tac.get().strip(),
            "chi_nhanh_bank_doi_tac": self.chi_nhanh_bank_doi_tac.get().strip(),
            "danh_sach_cv": self.jobs,
            "bang_cv_phang": bang_cv_phang,
            "tong_tien_chua_vat": format_money(tong_tien_chua_vat),
            "VAT": format_money(vat),
            "tong_tien_co_vat": format_money(tong_tien_co_vat),
            "tien_da_thanh_toan": format_money(tien_da_thanh_toan),
            "tien_con_phai_thanh_toan": format_money(tien_con_phai_thanh_toan),
            "tong_tien_co_vat_text": number_to_vietnamese_words(tong_tien_co_vat),
            "tien_con_phai_thanh_toan_text": number_to_vietnamese_words(
                tien_con_phai_thanh_toan
            ),
        }

        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)

        safe_ma_nt = safe_filename(self.ma_nt.get().strip() or "khong_so")
        safe_ma_tl = safe_filename(self.ma_tl.get().strip() or "khong_so")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # noqa: DTZ005

        filename = f"Bien_Ban_Nghiem_Thu_DV_{safe_ma_nt}_{timestamp}.docx"
        filename_tl = f"Bien_Ban_Thanh_Ly_{safe_ma_tl}_{timestamp}.docx"

        output_path = generate_docx(
            self.template_file,
            context,
            filename,
        )
        output_path_tl = generate_docx(
            self.template_file_thanh_li,
            context,
            filename_tl,
        )

        messagebox.showinfo(
            "Thành công",
            "Đã xuất thành công:\n\n"
            f"Biên bản nghiệm thu:\n{output_path}\n\n"
            f"Biên bản thanh lý:\n{output_path_tl}",
            parent=self,
        )

        open_file(output_path)
        open_file(output_path_tl)

    # ========================================================
    # EXPORT SALE
    # ========================================================

    def export_sale_docx(self):
        self.validate_common()

        if not self.ten_bbnt.get().strip():
            raise ValueError("Vui lòng nhập Tên biên bản.")

        if not self.hang_muc_nt_hdmb.get().strip():
            raise ValueError("Vui lòng nhập Hạng mục nghiệm thu.")

        if not os.path.exists(self.template_file):
            raise FileNotFoundError(
                "Không tìm thấy template mua bán:\n"
                + os.path.abspath(self.template_file)
            )

        if not self.products and not messagebox.askyesno(
            "Xác nhận",
            "Danh sách sản phẩm đang trống.\nBạn vẫn muốn xuất Word?",
            parent=self,
        ):
            return

        ngay_nt = self.ngay_nt.get().strip()
        ngay_hd = self.ngay_hd.get().strip()

        date_obj_nt = parse_date(ngay_nt, "Ngày nghiệm thu")
        date_obj_hd = parse_date(ngay_hd, "Ngày hợp đồng")

        max_len = max(
            len(self.partner_group),
            len(self.my_group),
        )

        member_rows = []
        for i in range(max_len):
            member_rows.append(
                {
                    "member": self.partner_group[i]
                    if i < len(self.partner_group)
                    else None,
                    "my_member": self.my_group[i] if i < len(self.my_group) else None,
                }
            )

        # Cập nhật chênh lệch theo công thức:
        # thực tế - hợp đồng.
        for product in self.products:
            product["chenh_lech"] = self.calculate_difference(
                product["sl_hd"],
                product["sl_thucte"],
            )

        context = {
            "date_nt": date_obj_nt.day,
            "month_nt": date_obj_nt.month,
            "year_nt": date_obj_nt.year,
            "date_hd": date_obj_hd.day,
            "month_hd": date_obj_hd.month,
            "year_hd": date_obj_hd.year,
            "ngay_hd": ngay_hd,
            "ngay_nt": ngay_nt,
            "ten_doi_tac": self.ten_doi_tac.get().strip(),
            "ma_hdmb": self.ma_hd.get().strip(),
            "ten_bbnt": self.ten_bbnt.get().strip(),
            "hang_muc_nt_hdmb": self.hang_muc_nt_hdmb.get().strip(),
            "gio_bat_dau_nt": self.gio_bat_dau_nt.get().strip(),
            "gio_ket_thuc_nt": self.gio_ket_thuc_nt.get().strip(),
            "muc_dich_hdmb": self.muc_dich_hdmb.get().strip(),
            "my_group": self.my_group,
            "partner_group": self.partner_group,
            "member_rows": member_rows,
            "products": self.products,
        }

        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)

        safe_ma_hd = safe_filename(self.ma_hd.get().strip() or "khong_so")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # noqa: DTZ005

        filename = f"Bien_Ban_Nghiem_Thu_MB_{safe_ma_hd}_{timestamp}.docx"

        output_path = generate_docx(
            self.template_file,
            context,
            filename,
        )

        messagebox.showinfo(
            "Thành công",
            f"Đã xuất file:\n{output_path}",
            parent=self,
        )

        open_file(output_path)


# ============================================================
# JOB DIALOG
# ============================================================


class JobDialog(tk.Toplevel):
    def __init__(self, parent, title, data=None):
        super().__init__(parent)

        self.title(title)
        self.geometry("500x420")
        self.resizable(False, False)
        self.result = None

        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text="Tên dịch vụ / công việc:").pack(
            anchor="w", padx=10, pady=(10, 3)
        )

        self.name_entry = ttk.Entry(self)
        self.name_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Công việc thực hiện:").pack(
            anchor="w", padx=10, pady=(10, 3)
        )

        self.work_entry = ttk.Entry(self)
        self.work_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Đơn vị:").pack(anchor="w", padx=10, pady=(10, 3))

        self.unit_entry = ttk.Entry(self)
        self.unit_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Số lượng:").pack(anchor="w", padx=10, pady=(10, 3))

        self.quantity_entry = ttk.Entry(self)
        self.quantity_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Đơn giá (VND):").pack(anchor="w", padx=10, pady=(10, 3))

        self.price_entry = ttk.Entry(self)
        self.price_entry.pack(fill="x", padx=10)

        if data:
            self.name_entry.insert(0, data.get("ten_dich_vu", ""))
            self.work_entry.insert(0, data.get("cong_viec_thuc_hien", ""))
            self.unit_entry.insert(0, data.get("don_vi", "Gói"))
            self.quantity_entry.insert(0, data.get("so_luong", 1))
            self.price_entry.insert(0, data.get("don_gia", 0))
        else:
            self.unit_entry.insert(0, "Gói")
            self.quantity_entry.insert(0, "1")
            self.price_entry.insert(0, "0")

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=20)

        ttk.Button(
            button_frame,
            text="Lưu",
            style="NTSuccess.TButton",
            command=self.save,
        ).pack(side="left", padx=5)

        ttk.Button(
            button_frame,
            text="Hủy",
            style="NTSecondary.TButton",
            command=self.destroy,
        ).pack(side="left", padx=5)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.name_entry.focus()
        self.wait_window()

    def save(self):
        name = self.name_entry.get().strip()
        work = self.work_entry.get().strip()
        unit = self.unit_entry.get().strip()

        if not name:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng nhập tên công việc.",
                parent=self,
            )
            return

        if not unit:
            unit = "Gói"

        try:
            quantity = parse_number(self.quantity_entry.get())
            price = parse_number(self.price_entry.get())
        except ValueError as exc:
            messagebox.showwarning("Thông báo", str(exc), parent=self)
            return

        if quantity < 0:
            messagebox.showwarning(
                "Thông báo",
                "Số lượng không được âm.",
                parent=self,
            )
            return

        if price < 0:
            messagebox.showwarning(
                "Thông báo",
                "Đơn giá không được âm.",
                parent=self,
            )
            return

        self.result = {
            "ten_dich_vu": name,
            "cong_viec_thuc_hien": work,
            "don_vi": unit,
            "so_luong": quantity,
            "don_gia": price,
            "thanh_tien": quantity * price,
        }
        self.destroy()


# ============================================================
# MEMBER DIALOG
# ============================================================


class MemberDialog(tk.Toplevel):
    def __init__(self, parent, title, data=None):
        super().__init__(parent)

        self.title(title)
        self.geometry("450x220")
        self.resizable(False, False)
        self.result = None

        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text="Họ và tên:").pack(anchor="w", padx=10, pady=(15, 3))

        self.name_entry = ttk.Entry(self)
        self.name_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Chức vụ:").pack(anchor="w", padx=10, pady=(10, 3))

        self.role_entry = ttk.Entry(self)
        self.role_entry.pack(fill="x", padx=10)

        if data:
            self.name_entry.insert(0, data.get("name", ""))
            self.role_entry.insert(0, data.get("role", ""))

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=20)

        ttk.Button(
            button_frame,
            text="Lưu",
            style="NTSuccess.TButton",
            command=self.save,
        ).pack(side="left", padx=5)

        ttk.Button(
            button_frame,
            text="Hủy",
            style="NTSecondary.TButton",
            command=self.destroy,
        ).pack(side="left", padx=5)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.name_entry.focus()
        self.wait_window()

    def save(self):
        name = self.name_entry.get().strip()
        role = self.role_entry.get().strip()

        if not name:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng nhập họ tên.",
                parent=self,
            )
            return

        self.result = {
            "name": name,
            "role": role,
        }
        self.destroy()


# ============================================================
# PRODUCT DIALOG
# ============================================================


class ProductDialog(tk.Toplevel):
    def __init__(self, parent, title, data=None):
        super().__init__(parent)

        self.title(title)
        self.geometry("500x350")
        self.resizable(False, False)
        self.result = None

        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text="Nội dung:").pack(anchor="w", padx=10, pady=(15, 3))

        self.name_entry = ttk.Entry(self)
        self.name_entry.pack(fill="x", padx=10)

        ttk.Label(self, text="Đơn vị tính:").pack(anchor="w", padx=10, pady=(10, 3))

        self.unit_entry = ttk.Entry(self)
        self.unit_entry.pack(fill="x", padx=10)

        ttk.Label(
            self,
            text="Số lượng theo hợp đồng:",
        ).pack(anchor="w", padx=10, pady=(10, 3))

        self.sl_hd_entry = ttk.Entry(self)
        self.sl_hd_entry.pack(fill="x", padx=10)

        ttk.Label(
            self,
            text="Số lượng thực tế:",
        ).pack(anchor="w", padx=10, pady=(10, 3))

        self.sl_thucte_entry = ttk.Entry(self)
        self.sl_thucte_entry.pack(fill="x", padx=10)

        if data:
            self.name_entry.insert(0, data.get("name", ""))
            self.unit_entry.insert(0, data.get("unit", ""))
            self.sl_hd_entry.insert(0, data.get("sl_hd", 0))
            self.sl_thucte_entry.insert(0, data.get("sl_thucte", 0))
        else:
            self.unit_entry.insert(0, "Cái")
            self.sl_hd_entry.insert(0, "0")
            self.sl_thucte_entry.insert(0, "0")

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=20)

        ttk.Button(
            button_frame,
            text="Lưu",
            style="NTSuccess.TButton",
            command=self.save,
        ).pack(side="left", padx=5)

        ttk.Button(
            button_frame,
            text="Hủy",
            style="NTSecondary.TButton",
            command=self.destroy,
        ).pack(side="left", padx=5)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.name_entry.focus()
        self.wait_window()

    def save(self):
        name = self.name_entry.get().strip()
        unit = self.unit_entry.get().strip() or "Cái"

        if not name:
            messagebox.showwarning(
                "Thông báo",
                "Vui lòng nhập nội dung.",
                parent=self,
            )
            return

        try:
            sl_hd = parse_number(self.sl_hd_entry.get())
            sl_thucte = parse_number(self.sl_thucte_entry.get())
        except ValueError as exc:
            messagebox.showwarning("Thông báo", str(exc), parent=self)
            return

        if sl_hd < 0 or sl_thucte < 0:
            messagebox.showwarning(
                "Thông báo",
                "Số lượng không được âm.",
                parent=self,
            )
            return

        self.result = {
            "name": name,
            "unit": unit,
            "sl_hd": sl_hd,
            "sl_thucte": sl_thucte,
            "chenh_lech": sl_thucte - sl_hd,
        }
        self.destroy()


# ============================================================
# STANDALONE TEST
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Quản lý biên bản nghiệm thu hợp đồng")
    root.geometry("1200x850")
    root.minsize(1000, 650)

    # NghiemThuFrame tự quản lý canvas + scrollbar.
    # Không bọc thêm một canvas bên ngoài.
    app = NghiemThuFrame(root)
    app.pack(
        fill="both",
        expand=True,
        padx=0,
        pady=0,
    )

    root.mainloop()
