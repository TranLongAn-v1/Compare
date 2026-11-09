import os
import re

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_HBA1C_THIEU_DAI_THAO_DUONG = "HbA1c phải kèm theo bệnh đái tháo đường"
COT_HIEN_THI = ["MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "MA_BENH", "MA_BENH_YHCT", "LOI"]
COT_BAT_BUOC = ["MA_NHOM", "TEN_DICH_VU", "MA_BA", "HO_TEN", "MA_THE", "MA_BENH", "MA_BENH_YHCT"]
MA_BENH_DAI_THAO_DUONG = ("E10", "E11", "E12", "E13", "E14")


def _chuan_hoa_cot(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _chuan_hoa_ma_nhom(val):
    if pd.isna(val):
        return None
    raw = str(val).strip().replace(",", ".")
    if not raw:
        return None
    try:
        return int(float(raw))
    except Exception:
        return None


def _chuan_hoa_ten_dich_vu(val):
    raw = "" if pd.isna(val) else str(val)
    return " ".join(raw.split()).casefold()


def _co_ma_benh_dai_thao_duong(*values):
    for val in values:
        if pd.isna(val):
            continue
        raw = str(val).upper()
        tokens = re.findall(r"[A-Z]\d{2}(?:[.\dA-Z]*)?", raw)
        if any(token.startswith(MA_BENH_DAI_THAO_DUONG) for token in tokens):
            return True
    return False


def chuan_doan_xml3_phong_xet_nghiem(path_file, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML3...", "blue")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File XML3 không có dữ liệu.")

    df = _chuan_hoa_cot(df)
    thieu_cot = [cot for cot in COT_BAT_BUOC if cot not in df.columns]
    if thieu_cot:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu_cot)}")

    if callback_status:
        callback_status("Đang lọc MA_NHOM = 1 và dịch vụ HbA1c...", "blue")

    ma_nhom = df["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    ten_dich_vu = df["TEN_DICH_VU"].apply(_chuan_hoa_ten_dich_vu)
    mask_hba1c = (ma_nhom == 1) & ten_dich_vu.str.contains("định lượng hba1c", regex=False)
    df_loc = df.loc[mask_hba1c, COT_BAT_BUOC].copy()

    if callback_status:
        callback_status("Đang kiểm tra mã bệnh E10-E14...", "blue")

    for cot in COT_BAT_BUOC:
        df_loc[cot] = df_loc[cot].fillna("").astype(str).str.strip()

    df_loc["LOI"] = df_loc.apply(
        lambda row: ""
        if _co_ma_benh_dai_thao_duong(row["MA_BENH"], row["MA_BENH_YHCT"])
        else LOI_HBA1C_THIEU_DAI_THAO_DUONG,
        axis=1,
    )
    df_loc["_UU_TIEN_LOI"] = (df_loc["LOI"] != "").astype(int)
    df_ket_qua = (
        df_loc.sort_values("_UU_TIEN_LOI", ascending=False, kind="stable")
        .drop(columns=["_UU_TIEN_LOI"])
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        so_loi = int((df_ket_qua["LOI"] == LOI_HBA1C_THIEU_DAI_THAO_DUONG).sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_phong_xet_nghiem(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="ChuanDoan_XML3", index=False)
        ws = writer.sheets["ChuanDoan_XML3"]
        fill_loi = PatternFill(fill_type="solid", fgColor="FFF59D")

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU", "LOI"}:
                width = 36 if ten_cot != "LOI" else 44
            elif ten_cot in {"MA_BENH", "MA_BENH_YHCT"}:
                width = 24
            else:
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                loi_value = row[loi_col_idx - 1].value
                if loi_value == LOI_HBA1C_THIEU_DAI_THAO_DUONG:
                    for cell in row:
                        cell.fill = fill_loi

    return file_xuat
