import os
import unicodedata
from datetime import date

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_NGUOI_LON_KE_DICH_VU_NHI = "Người lớn nhưng kê dịch vụ dưới 16 tuổi"
LOI_TRE_EM_KE_DICH_VU_NGUOI_LON = "Trẻ dưới 16 tuổi nhưng kê dịch vụ người lớn"
COT_HIEN_THI = [
    "MA_LK",
    "MA_BA",
    "MA_BN",
    "HO_TEN",
    "NGAY_SINH",
    "MA_THE",
    "MA_DICH_VU",
    "TEN_DICH_VU",
    "LOI",
]
COT_XML1_BAT_BUOC = ["MA_LK", "MA_BN", "NGAY_SINH"]
COT_XML3_BAT_BUOC = ["MA_LK", "MA_BA", "HO_TEN", "MA_THE", "MA_NHOM", "MA_DICH_VU", "TEN_DICH_VU"]
COT_PTTT_BAT_BUOC = ["MANHOM_BHYT", "MADICHVU", "TENDICHVU", "TENCHUYENKHOA"]


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


def _chuan_hoa_text(val):
    raw = _chuan_hoa_gia_tri(val)
    raw = unicodedata.normalize("NFC", raw)
    return " ".join(raw.split()).casefold()


def _chuan_hoa_ma_nhom(val):
    raw = _chuan_hoa_gia_tri(val).replace(",", ".")
    if not raw:
        return ""
    try:
        return str(int(float(raw)))
    except Exception:
        return raw


def _lay_12_so_ngay(val):
    raw = _chuan_hoa_gia_tri(val)
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:12] if len(digits) >= 12 else ""


def _lay_ngay(val):
    digits = _lay_12_so_ngay(val)
    if len(digits) != 12:
        return None
    try:
        return date(int(digits[0:4]), int(digits[4:6]), int(digits[6:8]))
    except ValueError:
        return None


def _dinh_dang_ngay_sinh(val):
    digits = _lay_12_so_ngay(val)
    if len(digits) != 12:
        return _chuan_hoa_gia_tri(val)
    return f"{digits[6:8]}/{digits[4:6]}/{digits[0:4]}"


def _tinh_tuoi(ngay_sinh, ngay_tham_chieu):
    if ngay_sinh is None or ngay_tham_chieu is None:
        return None
    tuoi = ngay_tham_chieu.year - ngay_sinh.year
    if (ngay_tham_chieu.month, ngay_tham_chieu.day) < (ngay_sinh.month, ngay_sinh.day):
        tuoi -= 1
    return tuoi


def _ngay_tham_chieu_xml3(row):
    if "NGAY_TH_YL" in row and _lay_ngay(row["NGAY_TH_YL"]):
        return _lay_ngay(row["NGAY_TH_YL"])
    if "NGAY_YL" in row and _lay_ngay(row["NGAY_YL"]):
        return _lay_ngay(row["NGAY_YL"])
    return date.today()


def _them_key_pttt(df, cot_ma_nhom, cot_ma_dich_vu, cot_ten_dich_vu):
    df = df.copy()
    df["_KEY_PTTT"] = (
        df[cot_ma_nhom].apply(_chuan_hoa_ma_nhom)
        + "|"
        + df[cot_ma_dich_vu].apply(_chuan_hoa_text)
        + "|"
        + df[cot_ten_dich_vu].apply(_chuan_hoa_text)
    )
    return df


def chuan_doan_xml3_pttt_tuoi(file_xml3, file_xml1, file_pttt, callback_status=None):
    if callback_status:
        callback_status("Đang đọc XML3, XML1 và danh mục PTTT...", "blue")

    df_xml3 = doc_file_thong_minh(file_xml3)
    df_xml1 = doc_file_thong_minh(file_xml1)
    df_pttt = doc_file_thong_minh(file_pttt)
    if df_xml3 is None or df_xml3.empty:
        raise ValueError("File XML3 không có dữ liệu.")
    if df_xml1 is None or df_xml1.empty:
        raise ValueError("File XML1 không có dữ liệu.")
    if df_pttt is None or df_pttt.empty:
        raise ValueError("File PTTT không có dữ liệu.")

    df_xml3 = _chuan_hoa_cot(df_xml3)
    df_xml1 = _chuan_hoa_cot(df_xml1)
    df_pttt = _chuan_hoa_cot(df_pttt)

    thieu_xml3 = [cot for cot in COT_XML3_BAT_BUOC if cot not in df_xml3.columns]
    thieu_xml1 = [cot for cot in COT_XML1_BAT_BUOC if cot not in df_xml1.columns]
    thieu_pttt = [cot for cot in COT_PTTT_BAT_BUOC if cot not in df_pttt.columns]
    if thieu_xml3:
        raise ValueError(f"XML3 thiếu cột bắt buộc: {', '.join(thieu_xml3)}")
    if thieu_xml1:
        raise ValueError(f"XML1 thiếu cột bắt buộc: {', '.join(thieu_xml1)}")
    if thieu_pttt:
        raise ValueError(f"PTTT thiếu cột bắt buộc: {', '.join(thieu_pttt)}")

    if callback_status:
        callback_status("Đang ghép NGAY_SINH từ XML1 vào XML3 theo MA_LK...", "blue")

    cot_ngay_xml3 = [cot for cot in ["NGAY_TH_YL", "NGAY_YL"] if cot in df_xml3.columns]
    cot_xml3_can_lay = COT_XML3_BAT_BUOC + cot_ngay_xml3
    df_xml3_work = df_xml3[cot_xml3_can_lay].copy()
    df_xml1_work = df_xml1[COT_XML1_BAT_BUOC].drop_duplicates(subset=["MA_LK"], keep="first").copy()

    for df in (df_xml3_work, df_xml1_work):
        df["MA_LK"] = df["MA_LK"].apply(_chuan_hoa_gia_tri)

    df_ghep = df_xml3_work.merge(df_xml1_work, on="MA_LK", how="left")

    if callback_status:
        callback_status("Đang đối chiếu danh mục PTTT theo MA_NHOM, MA_DICH_VU, TEN_DICH_VU...", "blue")

    df_ghep = _them_key_pttt(df_ghep, "MA_NHOM", "MA_DICH_VU", "TEN_DICH_VU")
    df_pttt = _them_key_pttt(df_pttt, "MANHOM_BHYT", "MADICHVU", "TENDICHVU")
    df_pttt["_LA_NHI_KHOA"] = df_pttt["TENCHUYENKHOA"].apply(_chuan_hoa_text) == "nhi khoa"
    df_pttt_key = df_pttt[["_KEY_PTTT", "_LA_NHI_KHOA"]].drop_duplicates(subset=["_KEY_PTTT"], keep="first")
    df_ghep = df_ghep.merge(df_pttt_key, on="_KEY_PTTT", how="inner")

    if callback_status:
        callback_status("Đang tính tuổi và kiểm tra kê sai dịch vụ theo tuổi...", "blue")

    for cot in ["MA_BA", "MA_BN", "HO_TEN", "MA_THE", "MA_DICH_VU", "TEN_DICH_VU", "NGAY_SINH"]:
        df_ghep[cot] = df_ghep[cot].fillna("").astype(str).str.strip()

    df_ghep["_NGAY_SINH_DATE"] = df_ghep["NGAY_SINH"].apply(_lay_ngay)
    df_ghep["_NGAY_THAM_CHIEU"] = df_ghep.apply(_ngay_tham_chieu_xml3, axis=1)
    df_ghep["_TUOI"] = df_ghep.apply(lambda row: _tinh_tuoi(row["_NGAY_SINH_DATE"], row["_NGAY_THAM_CHIEU"]), axis=1)
    df_ghep["LOI"] = ""

    la_tre_em = df_ghep["_TUOI"].notna() & (df_ghep["_TUOI"] < 16)
    la_nguoi_lon = df_ghep["_TUOI"].notna() & (df_ghep["_TUOI"] >= 16)
    la_nhi_khoa = df_ghep["_LA_NHI_KHOA"].fillna(False)
    df_ghep.loc[la_nguoi_lon & la_nhi_khoa, "LOI"] = LOI_NGUOI_LON_KE_DICH_VU_NHI
    df_ghep.loc[la_tre_em & ~la_nhi_khoa, "LOI"] = LOI_TRE_EM_KE_DICH_VU_NGUOI_LON

    df_ghep["NGAY_SINH"] = df_ghep["NGAY_SINH"].apply(_dinh_dang_ngay_sinh)
    df_ghep["_UU_TIEN_LOI"] = df_ghep["LOI"].ne("").astype(int)
    df_ket_qua = (
        df_ghep.sort_values("_UU_TIEN_LOI", ascending=False, kind="stable")
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        so_loi = int(df_ket_qua["LOI"].ne("").sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng PTTT, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_pttt_tuoi(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    fill_nguoi_lon = PatternFill(fill_type="solid", fgColor="D8B4FE")
    fill_tre_em = PatternFill(fill_type="solid", fgColor="FDBA74")

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="ChuanDoan_PTTT", index=False)
        ws = writer.sheets["ChuanDoan_PTTT"]

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU", "LOI"}:
                width = 34 if ten_cot != "LOI" else 44
            elif ten_cot in {"NGAY_SINH", "MA_DICH_VU"}:
                width = 18
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                loi = row[loi_col_idx - 1].value
                if loi == LOI_NGUOI_LON_KE_DICH_VU_NHI:
                    for cell in row:
                        cell.fill = fill_nguoi_lon
                elif loi == LOI_TRE_EM_KE_DICH_VU_NGUOI_LON:
                    for cell in row:
                        cell.fill = fill_tre_em

    return file_xuat
