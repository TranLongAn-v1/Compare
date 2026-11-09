import os
from datetime import datetime

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_THIEU_NGAY_GIUONG = "Thiếu ngày giường"
LOI_DU_NGAY_GIUONG = "Dư ngày giường"
MA_NHOM_GIUONG = {14, 15}
COT_HIEN_THI = [
    "MA_LK",
    "MA_BA",
    "MA_BN",
    "HO_TEN",
    "NGAY_VAO",
    "NGAY_VAO_NOI_TRU",
    "NGAY_RA",
    "MA_THE",
    "TEN_DICH_VU",
    "LOI",
]
COT_XML1_BAT_BUOC = ["MA_LK", "MA_BN", "NGAY_VAO", "NGAY_VAO_NOI_TRU", "NGAY_RA"]
COT_XML3_BAT_BUOC = ["MA_LK", "MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "MA_NHOM", "SO_LUONG"]


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
        return None
    try:
        return int(float(raw))
    except Exception:
        return None


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


def _chuyen_so_luong(val):
    raw = _chuan_hoa_gia_tri(val).replace(",", ".")
    if not raw:
        return 0.0
    try:
        return float(raw)
    except Exception:
        return 0.0


def _tinh_so_ngay_o_vien(ngay_vao, ngay_vao_noi_tru, ngay_ra):
    bat_dau = _parse_ngay_12_so(ngay_vao_noi_tru)
    if pd.isna(bat_dau):
        bat_dau = _parse_ngay_12_so(ngay_vao)
    ket_thuc = _parse_ngay_12_so(ngay_ra)
    if pd.isna(bat_dau) or pd.isna(ket_thuc) or ket_thuc < bat_dau:
        return None
    return (ket_thuc.date() - bat_dau.date()).days


def chuan_doan_xml3_ngay_giuong_noi_tru(file_xml3, file_xml1, callback_status=None):
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
        callback_status("Đang lọc ngày giường nội trú MA_NHOM 14 và 15...", "blue")

    df_xml3_work = df_xml3[COT_XML3_BAT_BUOC].copy()
    df_xml3_work["MA_NHOM_CHUAN"] = df_xml3_work["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    df_xml3_work = df_xml3_work[df_xml3_work["MA_NHOM_CHUAN"].isin(MA_NHOM_GIUONG)].copy()
    if df_xml3_work.empty:
        return pd.DataFrame(columns=COT_HIEN_THI)

    df_xml3_work["MA_LK"] = df_xml3_work["MA_LK"].apply(_chuan_hoa_gia_tri)
    df_xml3_work["SO_LUONG_CHUAN"] = df_xml3_work["SO_LUONG"].apply(_chuyen_so_luong)
    so_luong_theo_luot = df_xml3_work.groupby("MA_LK", dropna=False)["SO_LUONG_CHUAN"].sum().to_dict()

    df_xml1_work = df_xml1[COT_XML1_BAT_BUOC].copy()
    df_xml1_work["MA_LK"] = df_xml1_work["MA_LK"].apply(_chuan_hoa_gia_tri)
    df_xml1_work = df_xml1_work[df_xml1_work["MA_LK"] != ""].drop_duplicates(subset=["MA_LK"], keep="first")

    if callback_status:
        callback_status("Đang ghép ngày vào/ra từ XML1 theo MA_LK...", "blue")

    df_ghep = df_xml3_work.merge(df_xml1_work, on="MA_LK", how="inner")
    if df_ghep.empty:
        return pd.DataFrame(columns=COT_HIEN_THI)

    so_ngay_theo_luot = {}
    for _, row in df_ghep.drop_duplicates(subset=["MA_LK"], keep="first").iterrows():
        so_ngay_theo_luot[row["MA_LK"]] = _tinh_so_ngay_o_vien(
            row["NGAY_VAO"],
            row["NGAY_VAO_NOI_TRU"],
            row["NGAY_RA"],
        )

    def gan_loi(ma_lk):
        so_ngay = so_ngay_theo_luot.get(ma_lk)
        if so_ngay is None:
            return ""
        so_luong_giuong = so_luong_theo_luot.get(ma_lk, 0.0)
        if so_luong_giuong < so_ngay:
            return LOI_THIEU_NGAY_GIUONG
        if so_luong_giuong > so_ngay:
            return LOI_DU_NGAY_GIUONG
        return ""

    if callback_status:
        callback_status("Đang đối chiếu số lượng giường với số ngày ở viện...", "blue")

    df_ket_qua = df_ghep[COT_HIEN_THI[:-1]].copy()
    for cot in COT_HIEN_THI[:-1]:
        df_ket_qua[cot] = df_ket_qua[cot].apply(_chuan_hoa_gia_tri)
    df_ket_qua["LOI"] = df_ket_qua["MA_LK"].apply(gan_loi)
    for cot_ngay in ["NGAY_VAO", "NGAY_VAO_NOI_TRU", "NGAY_RA"]:
        df_ket_qua[cot_ngay] = df_ket_qua[cot_ngay].apply(_dinh_dang_ngay_hien_thi)

    df_ket_qua["_UU_TIEN_LOI"] = df_ket_qua["LOI"].ne("").astype(int)
    df_ket_qua = (
        df_ket_qua.sort_values(["_UU_TIEN_LOI", "MA_LK"], ascending=[False, True], kind="stable")
        .drop(columns=["_UU_TIEN_LOI"])
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        thieu = int(df_ket_qua["LOI"].eq(LOI_THIEU_NGAY_GIUONG).sum())
        du = int(df_ket_qua["LOI"].eq(LOI_DU_NGAY_GIUONG).sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, thiếu {thieu}, dư {du}.", "green")

    return df_ket_qua


def _merge_cac_cot_theo_ma_lk(ws, headers):
    if "MA_LK" not in headers:
        return
    ma_lk_col = headers.index("MA_LK") + 1
    ten_dich_vu_col = headers.index("TEN_DICH_VU") + 1 if "TEN_DICH_VU" in headers else None
    cot_can_gop = [idx for idx in range(1, len(headers) + 1) if idx != ten_dich_vu_col]
    row = 2
    max_row = ws.max_row
    while row <= max_row:
        ma_lk = ws.cell(row=row, column=ma_lk_col).value
        end_row = row
        while end_row + 1 <= max_row and ws.cell(row=end_row + 1, column=ma_lk_col).value == ma_lk:
            end_row += 1
        if ma_lk not in (None, "") and end_row > row:
            for col_idx in cot_can_gop:
                ws.merge_cells(start_row=row, start_column=col_idx, end_row=end_row, end_column=col_idx)
                ws.cell(row=row, column=col_idx).alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        row = end_row + 1


def xuat_excel_chuan_doan_xml3_ngay_giuong_noi_tru(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="NgayGiuongNoiTru", index=False)
        ws = writer.sheets["NgayGiuongNoiTru"]
        fill_thieu = PatternFill(fill_type="solid", fgColor="E9D5FF")
        fill_du = PatternFill(fill_type="solid", fgColor="FDBA74")

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU", "LOI"}:
                width = 34
            elif ten_cot in {"NGAY_VAO", "NGAY_VAO_NOI_TRU", "NGAY_RA"}:
                width = 20
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                loi = row[loi_col_idx - 1].value
                fill = None
                if loi == LOI_THIEU_NGAY_GIUONG:
                    fill = fill_thieu
                elif loi == LOI_DU_NGAY_GIUONG:
                    fill = fill_du
                if fill:
                    for cell in row:
                        cell.fill = fill
        _merge_cac_cot_theo_ma_lk(ws, headers)

    return file_xuat
