import os

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_TRUNG_DVKT = "Trùng dịch vụ kỹ thuật"
COT_HIEN_THI = ["MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_KQ", "NGAY_TH_YL", "LOI"]
COT_BAT_BUOC = ["MA_NHOM", "MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_KQ", "NGAY_TH_YL"]


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


def _chuan_hoa_ma_nhom(val):
    raw = _chuan_hoa_gia_tri(val).replace(",", ".")
    if not raw:
        return ""
    try:
        return str(int(float(raw)))
    except Exception:
        return raw


def _lay_12_so(val):
    raw = _chuan_hoa_gia_tri(val)
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:12] if len(digits) >= 12 else ""


def _dinh_dang_ngay_hien_thi(val):
    digits = _lay_12_so(val)
    if len(digits) != 12:
        return _chuan_hoa_gia_tri(val)
    return f"{digits[6:8]}/{digits[4:6]}/{digits[0:4]} {digits[8:10]}:{digits[10:12]}"


def chuan_doan_xml3_trung_dvkt(file_xml3, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML3...", "blue")

    df = doc_file_thong_minh(file_xml3)
    if df is None or df.empty:
        raise ValueError("File XML3 không có dữ liệu.")

    df = _chuan_hoa_cot(df)
    thieu_cot = [cot for cot in COT_BAT_BUOC if cot not in df.columns]
    if thieu_cot:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu_cot)}")

    if callback_status:
        callback_status("Đang kiểm tra trùng cặp NGAY_TH_YL và NGAY_KQ theo MA_BA, bỏ qua MA_NHOM = 1, 10...", "blue")

    df_work = df[COT_BAT_BUOC].copy()
    for cot in COT_BAT_BUOC:
        df_work[cot] = df_work[cot].apply(_chuan_hoa_gia_tri)

    ma_nhom = df_work["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    ngay_kq_key = df_work["NGAY_KQ"].apply(_lay_12_so)
    ngay_th_yl_key = df_work["NGAY_TH_YL"].apply(_lay_12_so)
    df_work["_NGAY_KQ_KEY"] = ngay_kq_key
    df_work["_NGAY_TH_YL_KEY"] = ngay_th_yl_key
    mask_xet = (~ma_nhom.isin({"1", "10"})) & (df_work["MA_BA"] != "") & (ngay_kq_key != "") & (ngay_th_yl_key != "")
    so_dong_trung = (
        df_work.loc[mask_xet]
        .groupby(["MA_BA", "_NGAY_TH_YL_KEY", "_NGAY_KQ_KEY"])["TEN_DICH_VU"]
        .transform("size")
    )
    mask_loi = pd.Series(False, index=df_work.index)
    mask_loi.loc[mask_xet] = so_dong_trung.gt(1).to_numpy()

    df_work["LOI"] = ""
    df_work.loc[mask_loi, "LOI"] = LOI_TRUNG_DVKT
    for cot_ngay in ["NGAY_KQ", "NGAY_TH_YL"]:
        df_work[cot_ngay] = df_work[cot_ngay].apply(_dinh_dang_ngay_hien_thi)

    df_work["_UU_TIEN_LOI"] = df_work["LOI"].ne("").astype(int)
    df_ket_qua = (
        df_work.sort_values("_UU_TIEN_LOI", ascending=False, kind="stable")
        .drop(columns=["_UU_TIEN_LOI", "MA_NHOM", "_NGAY_KQ_KEY", "_NGAY_TH_YL_KEY"])
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        so_loi = int(df_ket_qua["LOI"].eq(LOI_TRUNG_DVKT).sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_trung_dvkt(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="ChuanDoan_TrungDVKT", index=False)
        ws = writer.sheets["ChuanDoan_TrungDVKT"]
        fill_loi = PatternFill(fill_type="solid", fgColor="FFF59D")

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU", "LOI"}:
                width = 34 if ten_cot != "LOI" else 30
            elif ten_cot in {"NGAY_KQ", "NGAY_TH_YL"}:
                width = 18
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                if row[loi_col_idx - 1].value == LOI_TRUNG_DVKT:
                    for cell in row:
                        cell.fill = fill_loi

    return file_xuat
