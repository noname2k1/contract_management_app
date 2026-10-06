def flatten_jobs(job_list):
    """
    Chuyển danh sách công việc dạng cha/con
    thành danh sách phẳng để đưa vào docxtpl.

    Công việc cha:
        có children

    Công việc con:
        không có children
    """

    result = []

    for job in job_list:
        # ==========================================
        # CÔNG VIỆC CHA
        # ==========================================

        result.append(
            {
                "stt": job.get("stt", ""),
                "ten_dich_vu": job.get("ten_dich_vu", ""),
                "cong_viec_thuc_hien": job.get(
                    "cong_viec_thuc_hien",
                    "",
                ),
                "don_vi": job.get(
                    "don_vi",
                    "",
                ),
                "so_luong": job.get(
                    "so_luong",
                    0,
                ),
                "don_gia": job.get(
                    "don_gia",
                    0,
                ),
                "thanh_tien": job.get(
                    "thanh_tien",
                    0,
                ),
                "children": job.get(
                    "children",
                    [],
                ),
            }
        )

        # ==========================================
        # CÔNG VIỆC CON
        # ==========================================

        for child in job.get("children", []):
            result.append(
                {
                    "stt": child.get("stt", ""),
                    "ten_dich_vu": child.get(
                        "ten_dich_vu",
                        "",
                    ),
                    "cong_viec_thuc_hien": child.get(
                        "cong_viec_thuc_hien",
                        "",
                    ),
                }
            )

    return result


def calculate_job_total(job_list):
    """
    Tính tổng thành tiền của các công việc cha.

    Công việc con không được tính tiền.
    """

    total = 0

    for job in job_list:
        try:
            total += float(
                job.get(
                    "thanh_tien",
                    0,
                )
            )
        except (TypeError, ValueError):
            pass

    return total


def calculate_service_totals(
    jobs,
    vat_rate=8,
    tien_da_thanh_toan=0,
):
    """
    Tính toàn bộ giá trị dịch vụ.

    Trả về:

        tong_tien_chua_vat
        VAT
        tong_tien_co_vat
        tien_da_thanh_toan
        tien_con_phai_thanh_toan
    """

    tong_tien_chua_vat = calculate_job_total(jobs)

    vat = tong_tien_chua_vat * vat_rate / 100

    tong_tien_co_vat = tong_tien_chua_vat + vat

    tien_con_phai_thanh_toan = tong_tien_co_vat - tien_da_thanh_toan

    return {
        "tong_tien_chua_vat": tong_tien_chua_vat,
        "VAT": vat,
        "tong_tien_co_vat": tong_tien_co_vat,
        "tien_da_thanh_toan": tien_da_thanh_toan,
        "tien_con_phai_thanh_toan": tien_con_phai_thanh_toan,
    }
