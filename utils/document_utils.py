# utils/document_utils.py

import os

from docxtpl import DocxTemplate


def generate_docx(template_file, context_data, output_filename, output_dir="./outputs"):
    """
    Render template Word bằng docxtpl.

    Parameters
    ----------
    template_file : str
        Đường dẫn template .docx

    context_data : dict
        Dữ liệu truyền vào template

    output_filename : str
        Tên file Word đầu ra

    output_dir : str
        Thư mục chứa file xuất
    """

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"Không tìm thấy template:\n{template_file}")

    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(
        output_dir,
        output_filename,
    )

    doc = DocxTemplate(template_file)

    doc.render(context_data)

    doc.save(output_path)

    return output_path
