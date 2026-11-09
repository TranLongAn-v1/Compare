import os
from datetime import datetime

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_NGAY_KQ_LON_HON_NGAY_RA = "Ngày trả kết quả lớn hơn ngày ra"
COT_HIEN_THI = ["MA_BA", "MA_BN", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_VAO", "NGAY_KQ", "NGAY_RA", "LOI"]
COT_XML1_BAT_BUOC = ["MA_LK", "MA_BN", "NGAY_VAO", "NGAY_RA"]
COT_XML3_BAT_BUOC = ["MA_LK", "MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_KQ"]


def _chuan_hoa_cot(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _chuan_hoa_gia_tri(val):
    if pd.isna(val):
        return ""
    raw = str(val).strip()
    if raw.endswith(".0") and raw.replace(".", "", 1).isdigit():
        raw = raw[:-2]
    return raw


def _lay_12_so(val):
    raw = _chuan_hoa_gia_tri(val)
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:12] if len(digits) >= 12 else ""


def _parse_ngay_12_so(val):
    digits = _lay_12_so(val)
    if len(digits) != 12:
        return pd.NaT
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


def _dinh_dang_ngay_hien_thi(val):
    digits = _lay_12_so(val)
    if len(digits) != 12:
        return _chuan_hoa_gia_tri(val)
    return f"{digits[6:8]}/{digits[4:6]}/{digits[0:4]} {digits[8:10]}:{digits[10:12]}"


def chuan_doan_xml3_pttt_ngay_kq(file_xml3, file_xml1, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML3 và XML1...", "blue")

    df_xml3 = doc_file_thong_minh(file_xml3)
    df_xml1 = doc_file_thong_minh(file_xml1)
    if df_xml3 is None or df_xml3.empty:
        raise ValueError("File XML3 không có dữ liệu.")
    if df_xml1 is None or df_xml1.empty:
        raise ValueError("File XML1 không có dữ liệu.")

    df_xml3 = _chuan_hoa_cot(df_xml3)
    df_xml1 = _chuan_hoa_cot(df_xml1)
    thieu_xml3 = [cot for cot in COT_XML3_BAT_BUOC if cot not in df_xml3.columns]
    thieu_xml1 = [cot for cot in COT_XML1_BAT_BUOC if cot not in df_xml1.columns]
    if thieu_xml3:
        raise ValueError(f"XML3 thiếu cột bắt buộc: {', '.join(thieu_xml3)}")
    if thieu_xml1:
        raise ValueError(f"XML1 thiếu cột bắt buộc: {', '.join(thieu_xml1)}")

    if callback_status:
        callback_status("Đang ghép NGAY_VAO và NGAY_RA từ XML1 vào XML3 theo MA_LK...", "blue")

    df_xml3_work = df_xml3[COT_XML3_BAT_BUOC].copy()
    df_xml1_work = df_xml1[COT_XML1_BAT_BUOC].copy()
    df_xml3_work["MA_LK"] = df_xml3_work["MA_LK"].apply(_chuan_hoa_gia_tri)
    df_xml1_work["MA_LK"] = df_xml1_work["MA_LK"].apply(_chuan_hoa_gia_tri)
    df_xml1_work = df_xml1_work[df_xml1_work["MA_LK"] != ""].drop_duplicates(subset=["MA_LK"], keep="first")
    df_ghep = df_xml3_work.merge(df_xml1_work, on="MA_LK", how="inner")

    if callback_status:
        callback_status("Đang kiểm tra NGAY_KQ và NGAY_RA...", "blue")

    df_ket_qua = df_ghep[COT_HIEN_THI[:-1]].copy()
    for cot in COT_HIEN_THI[:-1]:
        df_ket_qua[cot] = df_ket_qua[cot].apply(_chuan_hoa_gia_tri)

    ngay_kq_dt = df_ket_qua["NGAY_KQ"].apply(_parse_ngay_12_so)
    ngay_ra_dt = df_ket_qua["NGAY_RA"].apply(_parse_ngay_12_so)
    mask_loi = ngay_kq_dt.notna() & ngay_ra_dt.notna() & (ngay_ra_dt < ngay_kq_dt)

    df_ket_qua["LOI"] = ""
    df_ket_qua.loc[mask_loi, "LOI"] = LOI_NGAY_KQ_LON_HON_NGAY_RA
    for cot_ngay in ["NGAY_VAO", "NGAY_KQ", "NGAY_RA"]:
        df_ket_qua[cot_ngay] = df_ket_qua[cot_ngay].apply(_dinh_dang_ngay_hien_thi)

    df_ket_qua["_UU_TIEN_LOI"] = df_ket_qua["LOI"].ne("").astype(int)
    df_ket_qua = (
        df_ket_qua.sort_values("_UU_TIEN_LOI", ascending=False, kind="stable")
        .drop(columns=["_UU_TIEN_LOI"])
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        so_loi = int(df_ket_qua["LOI"].eq(LOI_NGAY_KQ_LON_HON_NGAY_RA).sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_pttt_ngay_kq(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="ChuanDoan_NgayKQ", index=False)
        ws = writer.sheets["ChuanDoan_NgayKQ"]
        fill_loi = PatternFill(fill_type="solid", fgColor="FFF59D")

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU", "LOI"}:
                width = 34 if ten_cot != "LOI" else 36
            elif ten_cot in {"NGAY_VAO", "NGAY_KQ", "NGAY_RA"}:
                width = 18
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                if row[loi_col_idx - 1].value == LOI_NGAY_KQ_LON_HON_NGAY_RA:
                    for cell in row:
                        cell.fill = fill_loi

    return file_xuat
