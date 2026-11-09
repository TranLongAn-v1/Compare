import os
import re
import unicodedata
import warnings
import zlib
from decimal import Decimal, ROUND_HALF_UP
from numbers import Number

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh, lam_sach_chuoi

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)


CHENH_LECH_TOI_THIEU = Decimal("0.01")


def _chuan_hoa_ten_cot(val):
    raw = lam_sach_chuoi(val).lower()
    raw = "".join(ch for ch in unicodedata.normalize("NFD", raw) if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]", "", raw)


def _tim_cot(df, *ten_uu_tien):
    cot_chuan = {_chuan_hoa_ten_cot(c): c for c in df.columns}
    for ten in ten_uu_tien:
        key = _chuan_hoa_ten_cot(ten)
        if key in cot_chuan:
            return cot_chuan[key]
    raise KeyError(f"Khong tim thay cot: {', '.join(ten_uu_tien)}")


def _to_decimal(val):
    if pd.isna(val) or val == "":
        return Decimal("0")
    if isinstance(val, Number):
        return Decimal(str(val))
    text = str(val).strip().replace(" ", "")
    if not text:
        return Decimal("0")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    elif text.count(".") > 1:
        text = text.replace(".", "")
    elif text.count(".") == 1:
        before, after = text.split(".", 1)
        if len(after) == 3 and before.replace("-", "").isdigit():
            text = before + after
    try:
        return Decimal(text)
    except Exception:
        return Decimal("0")


def _fmt_tien(val):
    value = Decimal(val).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if value == value.to_integral():
        return f"{int(value):,}".replace(",", ".")
    whole, frac = f"{value:.2f}".split(".")
    return f"{int(whole):,}".replace(",", ".") + "," + frac


def _fmt_so_luot(val):
    return f"{int(Decimal(val)):,}".replace(",", ".")


def _doc_tong_mau01(file_paths, callback_status=None):
    tong = {
        "so_luot": Decimal("0"),
        "tong_chi_bv": Decimal("0"),
        "tong_chi_bh": Decimal("0"),
        "bao_hiem_tt": Decimal("0"),
        "benh_nhan_cct": Decimal("0"),
    }
    chi_tiet_file = []

    for path in file_paths:
        if callback_status:
            callback_status(f"Dang doc Mau01: {os.path.basename(path)}", "blue")
        df = doc_file_thong_minh(path).dropna(how="all")

        cot_tong_chi_bv = _tim_cot(df, "Tong chi BV", "Tổng chi BV", "T_TONGCHI_BV", "TONGCHIBV")
        cot_tong_chi_bh = _tim_cot(df, "Tong chi BH", "Tổng chi BH", "T_TONGCHI_BH", "TONGCHIBH")
        cot_bao_hiem_tt = _tim_cot(df, "Bao hiem TT", "Bảo hiểm TT", "T_BHTT", "TBHTT")
        cot_benh_nhan_cct = _tim_cot(df, "Benh nhan CCT", "Bệnh nhân CCT", "T_BNCCT", "BENHNHANCCT")

        so_luot = Decimal(len(df))
        tong_chi_bv = sum((_to_decimal(v) for v in df[cot_tong_chi_bv]), Decimal("0"))
        tong_chi_bh = sum((_to_decimal(v) for v in df[cot_tong_chi_bh]), Decimal("0"))
        bao_hiem_tt = sum((_to_decimal(v) for v in df[cot_bao_hiem_tt]), Decimal("0"))
        benh_nhan_cct = sum((_to_decimal(v) for v in df[cot_benh_nhan_cct]), Decimal("0"))

        tong["so_luot"] += so_luot
        tong["tong_chi_bv"] += tong_chi_bv
        tong["tong_chi_bh"] += tong_chi_bh
        tong["bao_hiem_tt"] += bao_hiem_tt
        tong["benh_nhan_cct"] += benh_nhan_cct

        chi_tiet_file.append(
            {
                "file": os.path.basename(path),
                "so_luot": so_luot,
                "tong_chi_bv": tong_chi_bv,
                "tong_chi_bh": tong_chi_bh,
                "bao_hiem_tt": bao_hiem_tt,
                "benh_nhan_cct": benh_nhan_cct,
            }
        )

    return tong, chi_tiet_file


def _giai_nen_pdf_streams(pdf_bytes):
    streams = []
    for match in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", pdf_bytes, re.S):
        stream = match.group(1)
        try:
            streams.append(zlib.decompress(stream))
        except Exception:
            streams.append(stream)
    return streams


def _doc_cmap(streams):
    cmap = {}
    for stream in streams:
        if b"beginbfrange" not in stream:
            continue
        text = stream.decode("latin1", errors="ignore")
        for start_hex, end_hex, unicode_hex in re.findall(
            r"<([0-9A-Fa-f]{4})><([0-9A-Fa-f]{4})><([0-9A-Fa-f]{4})>",
            text,
        ):
            start = int(start_hex, 16)
            end = int(end_hex, 16)
            unicode_start = int(unicode_hex, 16)
            for code in range(start, end + 1):
                cmap[code] = chr(unicode_start + code - start)
    return cmap


def _decode_pdf_text(raw, cmap):
    # PDF literal strings escape bytes, including bytes inside a two-byte glyph.
    decoded = bytearray()
    escapes = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}
    i = 0
    while i < len(raw):
        value = raw[i]
        i += 1
        if value == 0x5C and i < len(raw):
            value = raw[i]
            i += 1
            if value in (10, 13):
                if value == 13 and i < len(raw) and raw[i] == 10:
                    i += 1
                continue
            if ord("0") <= value <= ord("7"):
                octal = bytes([value])
                while len(octal) < 3 and i < len(raw) and ord("0") <= raw[i] <= ord("7"):
                    octal += bytes([raw[i]])
                    i += 1
                value = int(octal, 8) & 0xFF
            else:
                value = escapes.get(value, value)
        decoded.append(value)
    return "".join(cmap.get(decoded[i] * 256 + decoded[i + 1], "") for i in range(0, len(decoded) - 1, 2))


def _trich_text_pdf(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        data = pdf_file.read()
    streams = _giai_nen_pdf_streams(data)
    cmap = _doc_cmap(streams)
    if not cmap:
        raise ValueError("Khong doc duoc bang ma Unicode trong PDF.")

    items = []
    # Fonts are often set once for several lines; track each text matrix rather
    # than requiring a new Tf operator after every Tm.
    number = rb"[-+]?(?:\d+(?:\.\d*)?|\.\d+)"
    pattern = re.compile(
        rb"(?:" + number + rb"\s+){4}(?P<x>" + number + rb")\s+(?P<y>" + number + rb")\s+Tm"
        rb"|\((?P<text>(?:\\[\s\S]|[^\\()])*)\)\s*Tj"
    )
    for stream in streams:
        if b"Tj" not in stream or b" Tm" not in stream:
            continue
        position = None
        for match in pattern.finditer(stream):
            if match.group("x") is not None:
                position = (round(float(match.group("y")), 1), float(match.group("x")))
                continue
            if position is None:
                continue
            y, x = position
            text = _decode_pdf_text(match.group("text"), cmap).strip()
            if text:
                items.append((y, x, text))
    return items


def _parse_pdf_number(text):
    cleaned = re.sub(r"[^0-9,.\-]", "", text)
    if not cleaned:
        return Decimal("0")
    return _to_decimal(cleaned)


def _parse_pdf_count(text):
    cleaned = re.sub(r"[^0-9]", "", text)
    return Decimal(cleaned or "0")


def _doc_tong_mau02_pdf(pdf_path, callback_status=None):
    if callback_status:
        callback_status("Dang doc tong cong trong PDF Mau 02...", "blue")
    items = _trich_text_pdf(pdf_path)
    if not items:
        raise ValueError("Khong trich duoc noi dung tu PDF.")

    labels = [item for item in items if re.fullmatch(r"tong\s+cong(?:\s+a\s*\+\s*b)?\s*:?", _bo_dau(item[2]).casefold())]
    if len(labels) != 1:
        raise ValueError("Khong tim thay dong Tong cong A + B trong PDF.")
    y_tong, x_tong, _ = labels[0]
    # Labels and right-aligned amounts use different font sizes, so their
    # baselines can differ slightly within the same table row.
    dong_tong = sorted(
        [item for item in items if abs(item[0] - y_tong) <= 3 and item[1] > x_tong
         and re.fullmatch(r"-?\d+(?:[.,]\d+)*", item[2])],
        key=lambda item: item[1],
    )
    # Identify the numbered table columns (1)..(7), independent of page width.
    headers = []
    for y, x, text in items:
        match = re.match(r"^\(([1-9])\)", text)
        if match and y > y_tong:
            headers.append((y, x, int(match.group(1))))
    if not headers:
        raise ValueError("Khong doc duoc tieu de cot trong PDF Mau 02.")
    y_header = min(item[0] for item in headers)
    columns = sorted([item for item in headers if abs(item[0] - y_header) <= 3], key=lambda item: item[1])
    if ([item[2] for item in columns] != list(range(1, len(columns) + 1))
            or len(columns) < 5 or len(dong_tong) != len(columns)):
        raise ValueError("Khong doc du 5 chi tieu tren dong Tong cong trong PDF Mau 02. Vui long kiem tra bo cuc file.")
    return {
        "so_luot": _parse_pdf_count(dong_tong[0][2]),
        "tong_chi_bv": _parse_pdf_number(dong_tong[1][2]),
        "tong_chi_bh": _parse_pdf_number(dong_tong[2][2]),
        "bao_hiem_tt": _parse_pdf_number(dong_tong[3][2]),
        "benh_nhan_cct": _parse_pdf_number(dong_tong[4][2]),
    }


def _doc_tong_mau02_excel(file_path, callback_status=None):
    if callback_status:
        callback_status("Dang doc tong cong trong Excel Mau 02...", "blue")

    df = pd.read_excel(file_path, header=None, engine="openpyxl").dropna(how="all")
    dong_tong = None
    for _, row in df.iterrows():
        text = " ".join(lam_sach_chuoi(v) for v in row.dropna().tolist())
        if "Tong cong" in _bo_dau(text):
            dong_tong = row
            break
    if dong_tong is None:
        raise ValueError("Khong tim thay dong Tong cong A + B trong Excel Mau 02.")

    def value_at(col, parse_func=_parse_pdf_number):
        if col >= len(dong_tong):
            return Decimal("0")
        return parse_func(lam_sach_chuoi(dong_tong.iloc[col]))

    return {
        "so_luot": value_at(5, _parse_pdf_count),
        "tong_chi_bv": value_at(7),
        "tong_chi_bh": value_at(9),
        "bao_hiem_tt": value_at(11),
        "benh_nhan_cct": value_at(15),
    }


def _doc_tong_mau02(file_path, callback_status=None):
    ext = os.path.splitext(file_path)[1].casefold()
    if ext in {".xlsx", ".xls"}:
        return _doc_tong_mau02_excel(file_path, callback_status)
    return _doc_tong_mau02_pdf(file_path, callback_status)


def _bo_dau(text):
    raw = "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")
    return raw.replace("Đ", "D").replace("đ", "d").replace("Đ", "D").replace("Ä'", "d")


def _dong_ket_qua(chi_tieu, ten_mau01, ten_mau02, mau01, mau02, la_so_luot=False):
    lech = Decimal(mau01) - Decimal(mau02)
    khop = abs(lech) < (Decimal("1") if la_so_luot else CHENH_LECH_TOI_THIEU)
    fmt = _fmt_so_luot if la_so_luot else _fmt_tien
    return {
        "Chỉ tiêu": chi_tieu,
        "Cột Mẫu 01": ten_mau01,
        "Cột Mẫu 02": ten_mau02,
        "Tổng Mẫu 01": fmt(mau01),
        "Tổng Mẫu 02": fmt(mau02),
        "Lệch (Mẫu 01 - Mẫu 02)": fmt(lech),
        "Trạng thái": "Khớp" if khop else "Lệch",
        "color": "none" if khop else "red",
    }


def doi_chieu_01_bh_02_bh(file_mau01_list, file_pdf, callback_status=None):
    try:
        if not file_mau01_list:
            raise ValueError("Vui long chon it nhat 1 file Mau01.")
        if not file_pdf:
            raise ValueError("Vui long chon file Mau 02 rptTH.")

        tong_mau01, _chi_tiet_file = _doc_tong_mau01(file_mau01_list, callback_status)
        tong_mau02 = _doc_tong_mau02(file_pdf, callback_status)

        rows = [
            _dong_ket_qua("Tổng số bệnh nhân", "Số dòng bệnh nhân", "Số lượt", tong_mau01["so_luot"], tong_mau02["so_luot"], True),
            _dong_ket_qua("Tổng chi BV", "Tổng chi BV", "Tổng chi khám bệnh, chữa bệnh", tong_mau01["tong_chi_bv"], tong_mau02["tong_chi_bv"]),
            _dong_ket_qua("Tổng chi BH", "Tổng chi BH", "Chi phí khám bệnh, chữa bệnh bảo hiểm y tế", tong_mau01["tong_chi_bh"], tong_mau02["tong_chi_bh"]),
            _dong_ket_qua("Bảo hiểm TT", "Bảo hiểm TT", "Đề nghị quỹ bảo hiểm y tế thanh toán", tong_mau01["bao_hiem_tt"], tong_mau02["bao_hiem_tt"]),
            _dong_ket_qua("Bệnh nhân CCT", "Bệnh nhân CCT", "Người bệnh cùng chi trả", tong_mau01["benh_nhan_cct"], tong_mau02["benh_nhan_cct"]),
        ]
        df = pd.DataFrame(rows)
        df.insert(0, "STT", range(1, len(df) + 1))

        if callback_status:
            so_lech = int((df["Trạng thái"] == "Lệch").sum())
            callback_status(f"Hoan thanh. {so_lech} chi tieu lech.", "green" if so_lech == 0 else "red")
        return True, df
    except Exception as exc:
        if callback_status:
            callback_status("Loi khi doi chieu.", "red")
        return False, str(exc)


def xuat_excel_01_bh_02_bh(df_ket_qua, file_xuat):
    base, ext = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){ext}"
        counter += 1

    df_xuat = df_ket_qua.drop(columns=["color"], errors="ignore")
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_xuat.to_excel(writer, sheet_name="Doi_Chieu", index=False)
        ws = writer.sheets["Doi_Chieu"]
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        if "color" in df_ket_qua.columns:
            for row_idx, color in enumerate(df_ket_qua["color"], start=2):
                if color == "red":
                    for col_idx in range(1, len(df_xuat.columns) + 1):
                        ws.cell(row=row_idx, column=col_idx).fill = red_fill
        for col in ws.columns:
            header = col[0].value
            width = 12 if header == "STT" else 24
            if header in {"Chỉ tiêu", "Cột Mẫu 02"}:
                width = 42
            if header == "Trạng thái":
                width = 14
            ws.column_dimensions[col[0].column_letter].width = width
    return file_xuat
