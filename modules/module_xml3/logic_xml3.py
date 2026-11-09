import os

import pandas as pd

from core.utils import doc_file_thong_minh


COT_CAN_LAY_BAT_BUOC = ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY", "MA_BA", "HO_TEN"]
COT_CAN_LAY_BO_SUNG = ["SOPHIEU", "NGAY_TH_YL"]
MA_NHOM_LOAI_BO = {3, 8, 10, 13, 14, 15, 18}


def _chuan_hoa_ma_nhom(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    s = s.replace(",", ".")
    try:
        return int(float(s))
    except Exception:
        return None


def _dinh_dang_ngay_th_yl(val):
    if pd.isna(val):
        return ""
    raw = str(val).strip()
    if not raw:
        return ""
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) != 12:
        return raw

    nam = digits[0:4]
    thang = digits[4:6]
    ngay = digits[6:8]
    gio = digits[8:10]
    phut = digits[10:12]
    return f"{ngay}/{thang}/{nam}/{gio}:{phut}"


def tong_hop_xml3(file_xml3, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML3...", "blue")

    df = doc_file_thong_minh(file_xml3)
    if df is None or df.empty:
        raise ValueError("File XML3 không có dữ liệu.")

    thieu_cot = [c for c in COT_CAN_LAY_BAT_BUOC + ["MA_NHOM"] if c not in df.columns]
    if thieu_cot:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu_cot)}")

    if callback_status:
        callback_status("Đang lọc dữ liệu theo MA_NHOM...", "blue")

    ma_nhom_series = df["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    cot_ket_qua = COT_CAN_LAY_BAT_BUOC + COT_CAN_LAY_BO_SUNG
    for cot in COT_CAN_LAY_BO_SUNG:
        if cot not in df.columns:
            df[cot] = ""
    df_loc = df.loc[~ma_nhom_series.isin(MA_NHOM_LOAI_BO), cot_ket_qua].copy()

    for col in cot_ket_qua:
        df_loc[col] = df_loc[col].fillna("").astype(str)
    if "NGAY_TH_YL" in df_loc.columns:
        df_loc["NGAY_TH_YL"] = df_loc["NGAY_TH_YL"].apply(_dinh_dang_ngay_th_yl)

    if callback_status:
        callback_status(f"Tổng hợp xong: {len(df_loc)} dòng.", "green")

    return df_loc.reset_index(drop=True)


def xuat_excel_xml3(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="TongHop_XML3", index=False)
        worksheet = writer.sheets["TongHop_XML3"]

        for col in worksheet.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"MA_DICH_VU", "MA_MAY", "MA_BA"}:
                chieu_rong = 18
            elif ten_cot == "TEN_DICH_VU":
                chieu_rong = 40
            elif ten_cot == "HO_TEN":
                chieu_rong = 28
            else:
                chieu_rong = 20
            worksheet.column_dimensions[col[0].column_letter].width = chieu_rong

    return file_xuat
