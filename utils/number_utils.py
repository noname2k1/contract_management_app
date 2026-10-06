def parse_number(value):
    """
    Chuyển giá trị đầu vào thành int hoặc float.

    Ví dụ:
        "1000"       -> 1000
        "1.5"        -> 1.5
        "1,500"      -> 1500
        1000000      -> 1000000

    Không cho phép số âm.
    """
    value = str(value).strip()

    if not value:
        return 0

    # Cho phép nhập dạng 1,000 hoặc 1.000
    value = value.replace(",", "").replace(".", "")

    try:
        number = float(value)

        if number.is_integer():
            return abs(int(number))

        return abs(number)

    except ValueError:
        raise ValueError(f"Giá trị số không hợp lệ: {value}")


def format_money(value):
    """
    Format số tiền theo định dạng Việt Nam.

    Ví dụ:
        5000000      -> "5.000.000"
        61452000     -> "61.452.000"
        100000000.5  -> "100.000.001"
    """
    try:
        value = float(value)
        return f"{value:,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def number_to_vietnamese_words(number):
    """Chuyển số thành chữ tiếng Việt.
    Ví dụ:
        5000000
        -> "Năm triệu"
        61452000
        -> "Sáu mươi mốt triệu, bốn trăm năm mươi hai nghìn"
    """
    number = int(round(float(number)))
    if number == 0:
        return "Không"
    if number < 0:
        return "Âm " + number_to_vietnamese_words(abs(number))

    ones = [
        "không",
        "một",
        "hai",
        "ba",
        "bốn",
        "năm",
        "sáu",
        "bảy",
        "tám",
        "chín",
    ]

    units = [
        "",
        "nghìn",
        "triệu",
        "tỷ",
        "nghìn tỷ",
        "triệu tỷ",
    ]

    def read_three_digits(n, full=False):
        hundred = n // 100
        ten = (n % 100) // 10
        unit = n % 10

        result = []

        if hundred > 0 or full:
            result.append(ones[hundred])
            result.append("trăm")

        if ten > 1:
            result.append(ones[ten])
            result.append("mươi")

            if unit == 1:
                result.append("mốt")
            elif unit == 4:
                result.append("tư")
            elif unit == 5:
                result.append("lăm")
            elif unit > 0:
                result.append(ones[unit])

        elif ten == 1:
            result.append("mười")

            if unit == 5:
                result.append("lăm")
            elif unit > 0:
                result.append(ones[unit])

        elif unit > 0:
            if hundred > 0 or full:
                result.append("lẻ")

            if unit == 5 and (hundred > 0 or full):
                result.append("năm")
            else:
                result.append(ones[unit])

        return " ".join(result)

    groups = []

    while number > 0:
        groups.append(number % 1000)
        number //= 1000

    result = []

    for index in range(len(groups) - 1, -1, -1):
        group = groups[index]

        if group == 0:
            continue

        full = index < len(groups) - 1 and group < 100

        text = read_three_digits(group, full=full)

        if units[index]:
            text += " " + units[index]

        result.append(text)

    result_text = ", ".join(result)
    return result_text[0].upper() + result_text[1:]
