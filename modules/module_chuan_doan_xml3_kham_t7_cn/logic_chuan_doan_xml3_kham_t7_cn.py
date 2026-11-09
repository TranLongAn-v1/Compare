import os
import unicodedata
from datetime import date

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_KHAM_NGOAI_GIO = "Khám ngoài giờ"
LOI_NGOAI_GIO_HANH_CHINH = "Ngoài giờ hành chính"
TEN_DICH_VU_HOP_LE = "khám hồi sức cấp cứu và chống độc"
COT_HIEN_THI = ["MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_TH_YL", "LOI"]
COT_BAT_BUOC = ["MA_NHOM", "MA_BA", "HO_TEN", "MA_THE", "TEN_DICH_VU", "NGAY_TH_YL"]


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


def _chuan_hoa_text(val):
    raw = "" if pd.isna(val) else str(val)
    raw = unicodedata.normalize("NFC", raw)
    return " ".join(raw.split()).casefold()


def _lay_12_so_ngay_th_yl(val):
    if pd.isna(val):
        return ""
    raw = str(val).strip()
    if raw.endswith(".0"):
        raw = raw[:-2]
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:12] if len(digits) >= 12 else ""


def _la_thu_bay_chu_nhat(val):
    digits = _lay_12_so_ngay_th_yl(val)
    if len(digits) != 12:
        return False
    try:
        ngay = date(int(digits[0:4]), int(digits[4:6]), int(digits[6:8]))
    except ValueError:
        return False
    return ngay.weekday() >= 5


def _la_ngoai_gio_hanh_chinh(val):
    digits = _lay_12_so_ngay_th_yl(val)
    if len(digits) != 12:
        return False
    try:
        gio = int(digits[8:10])
        phut = int(digits[10:12])
    except ValueError:
        return False

    tong_phut = gio * 60 + phut
    ca_sang = 7 * 60 <= tong_phut <= 11 * 60 + 30
    ca_chieu = 13 * 60 + 30 <= tong_phut <= 17 * 60
    return not (ca_sang or ca_chieu)


def _dinh_dang_ngay_th_yl(val):
    digits = _lay_12_so_ngay_th_yl(val)
    if len(digits) != 12:
        return "" if pd.isna(val) else str(val).strip()
    return f"{digits[6:8]}/{digits[4:6]}/{digits[0:4]} {digits[8:10]}:{digits[10:12]}"


def chuan_doan_xml3_kham_t7_cn(path_file, callback_status=None):
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
        callback_status("Đang lọc MA_NHOM = 13...", "blue")

    ma_nhom = df["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    df_loc = df.loc[ma_nhom == 13, COT_BAT_BUOC].copy()

    if callback_status:
        callback_status("Đang kiểm tra thứ 7/chủ nhật và ngoài giờ hành chính...", "blue")

    for cot in COT_BAT_BUOC:
        df_loc[cot] = df_loc[cot].fillna("").astype(str).str.strip()

    ngay_cuoi_tuan = df_loc["NGAY_TH_YL"].apply(_la_thu_bay_chu_nhat)
    ngoai_gio_hanh_chinh = df_loc["NGAY_TH_YL"].apply(_la_ngoai_gio_hanh_chinh)
    ten_hop_le = df_loc["TEN_DICH_VU"].apply(_chuan_hoa_text) == TEN_DICH_VU_HOP_LE
    df_loc["LOI"] = ""
    df_loc.loc[ngay_cuoi_tuan & ~ten_hop_le, "LOI"] = LOI_KHAM_NGOAI_GIO
    df_loc.loc[ngoai_gio_hanh_chinh & ~ten_hop_le & (df_loc["LOI"] == ""), "LOI"] = LOI_NGOAI_GIO_HANH_CHINH
    df_loc["NGAY_TH_YL"] = df_loc["NGAY_TH_YL"].apply(_dinh_dang_ngay_th_yl)

    df_ket_qua = df_loc[COT_HIEN_THI].reset_index(drop=True)

    if callback_status:
        so_loi = int((df_ket_qua["LOI"] != "").sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_kham_t7_cn(df_ket_qua, file_xuat):
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
                width = 34 if ten_cot != "LOI" else 22
            elif ten_cot == "NGAY_TH_YL":
                width = 22
            else:
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                if row[loi_col_idx - 1].value:
                    for cell in row:
                        cell.fill = fill_loi

    return file_xuat
