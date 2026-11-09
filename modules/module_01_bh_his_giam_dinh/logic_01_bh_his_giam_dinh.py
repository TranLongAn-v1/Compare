import os
import re
import unicodedata
import warnings
from decimal import Decimal, ROUND_HALF_UP
from numbers import Number

import pandas as pd
from openpyxl.styles import Font
from openpyxl.styles import PatternFill

from core.utils import chuyen_doi_ngay_7980, doc_file_thong_minh, lam_sach_chuoi

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)

NGUONG_LECH = Decimal("0.01")
MONEY_KEYS = ["tong_chi_bv", "tong_chi_bh", "bhtt", "bncct"]
MONEY_LABELS = {
    "tong_chi_bv": ("T_TONGCHI_BV", "Tổng chi BV"),
    "tong_chi_bh": ("T_TONGCHI_BH", "Tổng chi BH"),
    "bhtt": ("T_BHTT", "Bảo hiểm TT"),
    "bncct": ("T_BNCCT", "Bệnh nhân CCT"),
}


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


def _fmt_so(val):
    return f"{int(Decimal(val)):,}".replace(",", ".")


def _bo_dau(text):
    raw = "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")
    return raw.replace("Đ", "D").replace("đ", "d")


def _chuan_hoa_ho_ten(val):
    raw = re.sub(r"\s+", " ", lam_sach_chuoi(val).upper()).strip()
    raw = _bo_dau(raw)
    return re.sub(r"[^A-Z0-9 ]", "", raw)


def _chuan_hoa_ma_the(val):
    raw = lam_sach_chuoi(val).upper()
    return re.sub(r"[^A-Z0-9]", "", raw)


def _parse_datetime(val):
    s = lam_sach_chuoi(val)
    if not s:
        return pd.NaT
    s = re.sub(r"\s+(SA|CH)\s*$", "", s, flags=re.IGNORECASE)
    if re.fullmatch(r"\d{12}", s):
        s = chuyen_doi_ngay_7980(s)
    dt = pd.to_datetime(s, format="%d/%m/%Y %H:%M", errors="coerce")
    if pd.notna(dt):
        return dt
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return dt
    return pd.NaT


def _fmt_datetime_key(dt):
    if pd.isna(dt):
        return ""
    return pd.Timestamp(dt).strftime("%Y%m%d%H%M")


def _fmt_datetime_display(key):
    if not key:
        return ""
    try:
        return pd.to_datetime(key, format="%Y%m%d%H%M").strftime("%d/%m/%Y %H:%M")
    except Exception:
        return key


def _empty_group():
    return {
        "ho_ten": "",
        "ma_the": "",
        "ngay_vao": "",
        "ngay_ra": "",
        "so_dong": Decimal("0"),
        "files": set(),
        **{key: Decimal("0") for key in MONEY_KEYS},
    }


def _add_to_group(groups, row):
    key = (row["ma_the_key"], row["ho_ten_key"], row["ngay_vao_key"], row["ngay_ra_key"])
    item = groups.setdefault(key, _empty_group())
    item["ho_ten"] = item["ho_ten"] or row["ho_ten"]
    item["ma_the"] = item["ma_the"] or row["ma_the"]
    item["ngay_vao"] = item["ngay_vao"] or row["ngay_vao_key"]
    item["ngay_ra"] = item["ngay_ra"] or row["ngay_ra_key"]
    item["so_dong"] += Decimal("1")
    if row.get("file"):
        item["files"].add(row["file"])
    for money_key in MONEY_KEYS:
        item[money_key] += row[money_key]


def _doc_his(file_his, callback_status=None):
    if callback_status:
        callback_status(f"Dang doc HIS: {os.path.basename(file_his)}", "blue")
    df = doc_file_thong_minh(file_his).dropna(how="all")
    cols = {
        "ho_ten": _tim_cot(df, "HO_TEN", "Ho ten", "Họ và tên"),
        "ma_the": _tim_cot(df, "MA_THE_BHYT", "MA_THE", "Mã thẻ"),
        "ngay_vao": _tim_cot(df, "NGAY_VAO", "Ngày vào"),
        "ngay_ra": _tim_cot(df, "NGAY_RA", "Ngày ra"),
        "tong_chi_bv": _tim_cot(df, "T_TONGCHI_BV", "Tong chi BV", "Tổng chi BV"),
        "tong_chi_bh": _tim_cot(df, "T_TONGCHI_BH", "Tong chi BH", "Tổng chi BH"),
        "bhtt": _tim_cot(df, "T_BHTT", "Bao hiem TT", "Bảo hiểm TT"),
        "bncct": _tim_cot(df, "T_BNCCT", "Benh nhan CCT", "Bệnh nhân CCT"),
    }
    rows = []
    for _, raw in df.iterrows():
        ngay_vao = _parse_datetime(raw[cols["ngay_vao"]])
        ngay_ra = _parse_datetime(raw[cols["ngay_ra"]])
        ho_ten = lam_sach_chuoi(raw[cols["ho_ten"]])
        ma_the = lam_sach_chuoi(raw[cols["ma_the"]])
        rows.append(
            {
                "ho_ten": ho_ten,
                "ma_the": ma_the,
                "ho_ten_key": _chuan_hoa_ho_ten(ho_ten),
                "ma_the_key": _chuan_hoa_ma_the(ma_the),
                "ngay_vao_key": _fmt_datetime_key(ngay_vao),
                "ngay_ra_key": _fmt_datetime_key(ngay_ra),
                "file": os.path.basename(file_his),
                **{key: _to_decimal(raw[cols[key]]) for key in MONEY_KEYS},
            }
        )
    return rows


def _doc_giam_dinh(file_list, callback_status=None):
    rows = []
    for path in file_list:
        if callback_status:
            callback_status(f"Dang doc Giam dinh: {os.path.basename(path)}", "blue")
        df = doc_file_thong_minh(path).dropna(how="all")
        cols = {
            "ho_ten": _tim_cot(df, "Họ và tên", "Ho ten", "HO_TEN"),
            "ma_the": _tim_cot(df, "Mã thẻ", "Ma the", "MA_THE", "MA_THE_BHYT"),
            "ngay_vao": _tim_cot(df, "Ngày vào", "Ngay vao", "NGAY_VAO"),
            "ngay_ra": _tim_cot(df, "Ngày ra", "Ngay ra", "NGAY_RA"),
            "tong_chi_bv": _tim_cot(df, "Tổng chi BV", "Tong chi BV", "T_TONGCHI_BV"),
            "tong_chi_bh": _tim_cot(df, "Tổng chi BH", "Tong chi BH", "T_TONGCHI_BH"),
            "bhtt": _tim_cot(df, "Bảo hiểm TT", "Bao hiem TT", "T_BHTT"),
            "bncct": _tim_cot(df, "Bệnh nhân CCT", "Benh nhan CCT", "T_BNCCT"),
        }
        for _, raw in df.iterrows():
            ngay_vao = _parse_datetime(raw[cols["ngay_vao"]])
            ngay_ra = _parse_datetime(raw[cols["ngay_ra"]])
            ho_ten = lam_sach_chuoi(raw[cols["ho_ten"]])
            ma_the = lam_sach_chuoi(raw[cols["ma_the"]])
            rows.append(
                {
                    "ho_ten": ho_ten,
                    "ma_the": ma_the,
                    "ho_ten_key": _chuan_hoa_ho_ten(ho_ten),
                    "ma_the_key": _chuan_hoa_ma_the(ma_the),
                    "ngay_vao_key": _fmt_datetime_key(ngay_vao),
                    "ngay_ra_key": _fmt_datetime_key(ngay_ra),
                    "file": os.path.basename(path),
                    **{key: _to_decimal(raw[cols[key]]) for key in MONEY_KEYS},
                }
            )
    return rows


def _gom_rows(rows):
    groups = {}
    for row in rows:
        _add_to_group(groups, row)
    return groups


def _tong_groups(groups):
    result = {"so_luot": sum((v["so_dong"] for v in groups.values()), Decimal("0"))}
    for key in MONEY_KEYS:
        result[key] = sum((v[key] for v in groups.values()), Decimal("0"))
    return result


def _tao_dong_tong(chi_tieu, his_val, gd_val, money_key=None):
    lech = Decimal(his_val) - Decimal(gd_val)
    khop = abs(lech) < (Decimal("1") if money_key is None else NGUONG_LECH)
    row = {
        "Loại": "Tổng",
        "Họ tên / Chỉ tiêu": chi_tieu,
        "Mã thẻ": "",
        "Ngày vào": "",
        "Ngày ra": "",
        "Số dòng HIS": "",
        "Số dòng giám định": "",
        "Nguồn giám định": "",
        "Trạng thái": "Khớp" if khop else "Lệch",
        "color": "none" if khop else "red",
    }
    for key in MONEY_KEYS:
        his_col, gd_col = MONEY_LABELS[key]
        row[f"{his_col} HIS"] = ""
        row[f"{gd_col} giám định"] = ""
        row[f"Lệch {his_col}"] = ""
    if money_key is None:
        row["Số dòng HIS"] = _fmt_so(his_val)
        row["Số dòng giám định"] = _fmt_so(gd_val)
    else:
        his_col, gd_col = MONEY_LABELS[money_key]
        row[f"{his_col} HIS"] = _fmt_tien(his_val)
        row[f"{gd_col} giám định"] = _fmt_tien(gd_val)
        row[f"Lệch {his_col}"] = _fmt_tien(lech)
    return row


def _tao_dong_chi_tiet(key, his, gd):
    missing_his = his is None
    missing_gd = gd is None
    his = his or _empty_group()
    gd = gd or _empty_group()
    lech_cols = []
    for money_key in MONEY_KEYS:
        his_col, _gd_col = MONEY_LABELS[money_key]
        if abs(his[money_key] - gd[money_key]) >= NGUONG_LECH:
            lech_cols.append(his_col)

    if missing_his:
        trang_thai = "Có trên giám định, thiếu HIS"
        mau_to = "blue"
    elif missing_gd:
        trang_thai = "Có trên HIS, thiếu giám định"
        mau_to = "yellow"
    else:
        trang_thai = "Lệch tiền: " + ", ".join(lech_cols)
        mau_to = "red"

    row = {
        "Loại": "Chi tiết",
        "Họ tên / Chỉ tiêu": his["ho_ten"] or gd["ho_ten"],
        "Mã thẻ": his["ma_the"] or gd["ma_the"],
        "Ngày vào": _fmt_datetime_display(his["ngay_vao"] or gd["ngay_vao"] or key[2]),
        "Ngày ra": _fmt_datetime_display(his["ngay_ra"] or gd["ngay_ra"] or key[3]),
        "Số dòng HIS": _fmt_so(his["so_dong"]) if his["so_dong"] else "0",
        "Số dòng giám định": _fmt_so(gd["so_dong"]) if gd["so_dong"] else "0",
        "Nguồn giám định": "; ".join(sorted(gd["files"])),
        "Trạng thái": trang_thai,
        "color": mau_to,
    }
    for money_key in MONEY_KEYS:
        his_col, gd_col = MONEY_LABELS[money_key]
        row[f"{his_col} HIS"] = _fmt_tien(his[money_key])
        row[f"{gd_col} giám định"] = _fmt_tien(gd[money_key])
        row[f"Lệch {his_col}"] = _fmt_tien(his[money_key] - gd[money_key])
    return row


def doi_chieu_01_bh_his_giam_dinh(file_his, file_giam_dinh_list, callback_status=None):
    try:
        if not file_his:
            raise ValueError("Vui long chon file 01_MAU_01BH tu HIS.")
        if not file_giam_dinh_list:
            raise ValueError("Vui long chon it nhat 1 file Mau01BH giam dinh.")

        his_groups = _gom_rows(_doc_his(file_his, callback_status))
        gd_groups = _gom_rows(_doc_giam_dinh(file_giam_dinh_list, callback_status))
        tong_his = _tong_groups(his_groups)
        tong_gd = _tong_groups(gd_groups)

        rows = [
            _tao_dong_tong("Tổng số bệnh nhân", tong_his["so_luot"], tong_gd["so_luot"]),
            _tao_dong_tong("T_TONGCHI_BV / Tổng chi BV", tong_his["tong_chi_bv"], tong_gd["tong_chi_bv"], "tong_chi_bv"),
            _tao_dong_tong("T_TONGCHI_BH / Tổng chi BH", tong_his["tong_chi_bh"], tong_gd["tong_chi_bh"], "tong_chi_bh"),
            _tao_dong_tong("T_BHTT / Bảo hiểm TT", tong_his["bhtt"], tong_gd["bhtt"], "bhtt"),
            _tao_dong_tong("T_BNCCT / Bệnh nhân CCT", tong_his["bncct"], tong_gd["bncct"], "bncct"),
        ]

        for key in sorted(set(his_groups) | set(gd_groups), key=lambda k: (k[1], k[2], k[3], k[0])):
            his = his_groups.get(key)
            gd = gd_groups.get(key)
            if his is None or gd is None:
                rows.append(_tao_dong_chi_tiet(key, his, gd))
                continue
            has_money_diff = any(abs(his[money_key] - gd[money_key]) >= NGUONG_LECH for money_key in MONEY_KEYS)
            has_count_diff = his["so_dong"] != gd["so_dong"]
            if has_money_diff or has_count_diff:
                rows.append(_tao_dong_chi_tiet(key, his, gd))

        df = pd.DataFrame(rows)
        df.insert(0, "STT", range(1, len(df) + 1))
        if callback_status:
            so_chi_tiet = int((df["Loại"] == "Chi tiết").sum())
            callback_status(f"Hoan thanh. {so_chi_tiet} dong chi tiet lech.", "green" if so_chi_tiet == 0 else "red")
        return True, df
    except Exception as exc:
        if callback_status:
            callback_status("Loi khi doi chieu.", "red")
        return False, str(exc)


def xuat_excel_01_bh_his_giam_dinh(df_ket_qua, file_xuat):
    base, ext = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){ext}"
        counter += 1

    df_xuat = df_ket_qua.drop(columns=["color"], errors="ignore")
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_xuat.to_excel(writer, sheet_name="Doi_Chieu", index=False)
        ws = writer.sheets["Doi_Chieu"]
        fills = {
            "red": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),
            "blue": PatternFill(start_color="B3D9FF", end_color="B3D9FF", fill_type="solid"),
        }
        if "color" in df_ket_qua.columns:
            for row_idx, color in enumerate(df_ket_qua["color"], start=2):
                fill = fills.get(color)
                if fill:
                    for col_idx in range(1, len(df_xuat.columns) + 1):
                        ws.cell(row=row_idx, column=col_idx).fill = fill
        highlight_font = Font(color="4C1D95", bold=True)
        for row_idx, row in df_xuat.iterrows():
            excel_row = row_idx + 2
            for col_idx, header in enumerate(df_xuat.columns):
                if not (str(header).startswith("L") and "T_" in str(header)):
                    continue
                if abs(_to_decimal(row.get(header, 0))) < NGUONG_LECH:
                    continue
                for target_idx in [col_idx - 2, col_idx - 1, col_idx]:
                    if target_idx >= 0:
                        ws.cell(row=excel_row, column=target_idx + 1).font = highlight_font
        for col in ws.columns:
            header = col[0].value
            width = 14
            if header in {"Họ tên / Chỉ tiêu", "Trạng thái", "Nguồn giám định"}:
                width = 34
            elif header in {"Mã thẻ", "Ngày vào", "Ngày ra"}:
                width = 20
            elif "HIS" in str(header) or "giám định" in str(header) or "Lệch" in str(header):
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width
    return file_xuat
