import os
from datetime import datetime

import pandas as pd

from core.utils import doc_file_thong_minh, lam_sach_chuoi


def _clean_digits(value):
    text = lam_sach_chuoi(value)
    if text.endswith(".0"):
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def _clean_hospital_code(value):
    text = lam_sach_chuoi(value)
    if text.endswith(".0"):
        text = text[:-2]
    if text.isdigit() and len(text) < 5:
        return text.zfill(5)
    return text


def _parse_yyyymmddhhmm(value):
    digits = _clean_digits(value)
    if len(digits) < 8:
        return pd.NaT
    digits = digits[:12].ljust(12, "0")
    try:
        return datetime(
            int(digits[0:4]),
            int(digits[4:6]),
            int(digits[6:8]),
            int(digits[8:10]),
            int(digits[10:12]),
        )
    except ValueError:
        return pd.NaT


def _format_datetime(value):
    dt = _parse_yyyymmddhhmm(value)
    if pd.isna(dt):
        return ""
    return dt.strftime("%d/%m/%Y %H:%M")


def _format_gender(value):
    text = lam_sach_chuoi(value)
    if text.endswith(".0"):
        text = text[:-2]
    if text == "1":
        return "Nam"
    if text == "2":
        return "Nữ"
    return text


def _has_value(value):
    text = lam_sach_chuoi(value)
    return bool(text and text.casefold() != "nan")


def _required_columns(df, columns):
    missing = [col for col in columns if col not in df.columns]
    if missing:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(missing)}")


def _format_days(start_value, end_value):
    start = _parse_yyyymmddhhmm(start_value)
    end = _parse_yyyymmddhhmm(end_value)
    if pd.isna(start) or pd.isna(end):
        return ""
    total_minutes = int(round((end - start).total_seconds() / 60))
    if total_minutes < 0:
        return ""
    days, remaining_minutes = divmod(total_minutes, 1440)
    hours, minutes = divmod(remaining_minutes, 60)
    parts = [f"{days} ngày"]
    if hours:
        parts.append(f"{hours} giờ")
    if minutes:
        parts.append(f"{minutes} phút")
    return " ".join(parts)


def _tim_cot(df, candidates):
    normalized = {str(col).strip().casefold(): col for col in df.columns}
    for candidate in candidates:
        key = candidate.strip().casefold()
        if key in normalized:
            return normalized[key]
    return None


def _doc_danh_muc_co_so_kcb(path_file_danh_muc):
    df = pd.read_excel(path_file_danh_muc, header=1, dtype=str)
    df.columns = [str(col).strip() for col in df.columns]

    col_ma = _tim_cot(df, ["ma_bv", "MA_BV", "Mã BV", "Mã bệnh viện"])
    col_ten = _tim_cot(df, ["ten_bv", "TEN_BV", "Tên bệnh viện", "Tên cơ sở KCB"])
    col_dia_chi = _tim_cot(df, ["dia_chi", "DIA_CHI", "Địa chỉ"])
    missing = []
    if not col_ma:
        missing.append("ma_bv")
    if not col_ten:
        missing.append("ten_bv")
    if not col_dia_chi:
        missing.append("dia_chi")
    if missing:
        raise ValueError(f"File danh mục cơ sở KCB thiếu cột: {', '.join(missing)}")

    df_map = df[[col_ma, col_ten, col_dia_chi]].copy()
    df_map.columns = ["Mã nơi đến", "Tên bệnh viện", "Địa chỉ nơi chuyển đến"]
    df_map["Mã nơi đến"] = df_map["Mã nơi đến"].apply(_clean_hospital_code)
    df_map["Tên bệnh viện"] = df_map["Tên bệnh viện"].apply(lam_sach_chuoi)
    df_map["Địa chỉ nơi chuyển đến"] = df_map["Địa chỉ nơi chuyển đến"].apply(lam_sach_chuoi)
    df_map = df_map[df_map["Mã nơi đến"] != ""]
    return df_map.drop_duplicates(subset=["Mã nơi đến"], keep="first")


def _base_detail(df, df_danh_muc):
    detail = pd.DataFrame(
        {
            "Họ tên": df["HO_TEN"].apply(lam_sach_chuoi),
            "HOTEN_NGUOI_HT": df["HOTEN_NGUOI_HT"].apply(lam_sach_chuoi),
            "Ngày sinh": df["NGAY_SINH"].apply(_format_datetime),
            "Giới tính": df["GIOI_TINH"].apply(_format_gender),
            "Mã thẻ BHYT": df["MA_THE_BHYT"].apply(lam_sach_chuoi),
            "Ngày chuyển": df["NGAY_RA"].apply(_format_datetime),
            "Nơi chuyển đến": df["MA_NOI_DEN"].apply(_clean_hospital_code),
            "Địa chỉ bệnh nhân": df["DIA_CHI"].apply(lam_sach_chuoi),
        }
    )
    detail = detail.merge(
        df_danh_muc,
        how="left",
        left_on="Nơi chuyển đến",
        right_on="Mã nơi đến",
    )
    detail = detail.drop(columns=["Mã nơi đến"])
    detail["Tên bệnh viện"] = detail["Tên bệnh viện"].fillna("")
    detail["Địa chỉ nơi chuyển đến"] = detail["Địa chỉ nơi chuyển đến"].fillna("")
    cols = list(detail.columns)
    insert_after = cols.index("Nơi chuyển đến") + 1
    for col in ["Địa chỉ nơi chuyển đến", "Tên bệnh viện"]:
        cols.remove(col)
    cols[insert_after:insert_after] = ["Tên bệnh viện", "Địa chỉ nơi chuyển đến"]
    return detail[cols]


def tong_hop_so_lieu_chuyen_vien_tuan(path_file, path_file_danh_muc, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML13...", "blue")
    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không có dữ liệu.")

    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    required = [
        "HO_TEN",
        "HOTEN_NGUOI_HT",
        "NGAY_SINH",
        "GIOI_TINH",
        "MA_THE_BHYT",
        "NGAY_RA",
        "MA_NOI_DEN",
        "DIA_CHI",
        "NGAY_VAO",
        "NGAY_VAO_NOI_TRU",
        "CHAN_DOAN_RV",
    ]
    _required_columns(df, required)

    if callback_status:
        callback_status("Đang đọc danh mục cơ sở KCB...", "blue")
    df_danh_muc = _doc_danh_muc_co_so_kcb(path_file_danh_muc)

    if callback_status:
        callback_status("Đang tách ngoại trú/nội trú...", "blue")
    mask_noi_tru = df["NGAY_VAO_NOI_TRU"].apply(_has_value)
    df_noi_tru = df.loc[mask_noi_tru].copy().reset_index(drop=True)
    df_ngoai_tru = df.loc[~mask_noi_tru].copy().reset_index(drop=True)

    df_tong_hop = pd.DataFrame(
        [
            {"Chỉ tiêu": "Tổng số bệnh nhân chuyển viện", "Số lượng": len(df)},
            {"Chỉ tiêu": "Bệnh nhân ngoại trú chuyển tuyến", "Số lượng": len(df_ngoai_tru)},
            {"Chỉ tiêu": "Bệnh nhân nội trú chuyển tuyến", "Số lượng": len(df_noi_tru)},
        ]
    )

    df_ct_ngoai_tru = _base_detail(df_ngoai_tru, df_danh_muc)
    df_ct_noi_tru = _base_detail(df_noi_tru, df_danh_muc)
    df_ct_noi_tru.insert(4, "Chẩn đoán", df_noi_tru["CHAN_DOAN_RV"].apply(lam_sach_chuoi))
    df_ct_noi_tru.insert(5, "Ngày vào", df_noi_tru["NGAY_VAO"].apply(_format_datetime))
    df_ct_noi_tru.insert(6, "Ngày vào nội trú", df_noi_tru["NGAY_VAO_NOI_TRU"].apply(_format_datetime))
    df_ct_noi_tru.insert(
        7,
        "Thời gian điều trị tại khoa",
        [
            _format_days(row["NGAY_VAO_NOI_TRU"], row["NGAY_RA"])
            for _, row in df_noi_tru.iterrows()
        ],
    )

    if callback_status:
        callback_status("Tổng hợp số liệu chuyển viện xong.", "green")
    return df_tong_hop, df_ct_ngoai_tru, df_ct_noi_tru


def xuat_excel_so_lieu_chuyen_vien_tuan(df_tong_hop, df_ngoai_tru, df_noi_tru, file_xuat):
    base, ext = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){ext}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_tong_hop.to_excel(writer, sheet_name="Tong_hop", index=False)
        df_ngoai_tru.to_excel(writer, sheet_name="Ngoai_tru_chuyen_tuyen", index=False)
        df_noi_tru.to_excel(writer, sheet_name="Noi_tru_chuyen_tuyen", index=False)

        for sheet_name, worksheet in writer.sheets.items():
            for col in worksheet.columns:
                header = str(col[0].value or "")
                if header in {"Họ tên", "HOTEN_NGUOI_HT", "Địa chỉ bệnh nhân", "Địa chỉ nơi chuyển đến", "Chẩn đoán"}:
                    width = 34
                elif header == "Tên bệnh viện":
                    width = 32
                elif header in {"Mã thẻ BHYT", "Ngày sinh", "Ngày chuyển", "Ngày vào", "Ngày vào nội trú"}:
                    width = 20
                elif header == "Thời gian điều trị tại khoa":
                    width = 26
                elif header == "Chỉ tiêu":
                    width = 36
                else:
                    width = 18
                worksheet.column_dimensions[col[0].column_letter].width = width

    return file_xuat
