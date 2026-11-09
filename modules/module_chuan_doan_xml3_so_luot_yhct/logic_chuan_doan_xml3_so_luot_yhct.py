import math
import os
import re
import unicodedata

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh


LOI_VUOT_GIOI_HAN = "Chỉ định DVKT vượt quá giới hạn cho phép trong ngày"
COT_HIEN_THI = ["NGUOI_THUC_HIEN", "MA_BAC_SI", "MA_BA", "TEN_DICH_VU", "NGAY_KQ", "NGAY_TH_YL", "LOI"]
COT_BAT_BUOC = ["MA_NHOM", "NGUOI_THUC_HIEN", "MA_BAC_SI", "MA_BA", "TEN_DICH_VU", "NGAY_KQ", "NGAY_TH_YL"]
COT_TIEN = ["DON_GIA_BV", "DON_GIA_BH", "THANH_TIEN_BV", "THANH_TIEN_BH", "T_BNCCT", "T_BHTT"]
COT_HIEN_THI = [
    "NGUOI_THUC_HIEN",
    "MA_BAC_SI",
    "MA_BA",
    "TEN_DICH_VU",
    *COT_TIEN,
    "NGAY_KQ",
    "NGAY_TH_YL",
    "LOI",
]
COT_BAT_BUOC = [*COT_BAT_BUOC, *COT_TIEN]

CAU_HINH_DICH_VU_MAC_DINH = [
    ("Điện châm / ôn châm / Thủy châm", 30, ["điện châm", "ôn châm", "thủy châm"]),
    ("Kéo giãn cột sống bằng máy", 30, ["kéo giãn cột sống"]),
    ("Tập vận động (thụ động, chủ động, có trợ giúp)", 30, ["tập vận động"]),
    ("Cứu ngải / Giác hơi", 30, ["cứu ngải", "giác hơi"]),
    ("Ngôn ngữ trị liệu", 40, ["ngôn ngữ trị liệu"]),
    ("Tập hoạt động trị liệu", 50, ["tập hoạt động trị liệu"]),
    ("Xoa bóp bấm huyệt", 30, ["xoa bóp bấm huyệt"]),
    ("Chiếu đèn hồng ngoại", 15, ["chiếu đèn hồng ngoại"]),
    ("Điều trị bằng siêu âm / sóng ngắn / vi sóng", 15, ["điều trị bằng siêu âm", "sóng ngắn", "vi sóng"]),
    ("Kéo nắn, nắn chỉnh khớp nhẹ nhàng", 15, ["kéo nắn", "nắn chỉnh khớp"]),
]


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
    raw = unicodedata.normalize("NFC", _chuan_hoa_gia_tri(val))
    return " ".join(raw.split()).casefold()


def _lay_12_so(val):
    raw = _chuan_hoa_gia_tri(val)
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:12] if len(digits) >= 12 else ""


def _chuan_hoa_ma_nhom(val):
    raw = _chuan_hoa_gia_tri(val).replace(",", ".")
    if not raw:
        return ""
    try:
        return str(int(float(raw)))
    except Exception:
        return raw


def _lay_ngay_key(row):
    for cot in ("NGAY_TH_YL", "NGAY_KQ"):
        digits = _lay_12_so(row.get(cot, ""))
        if len(digits) == 12:
            return digits[:8]
    return ""


def _dinh_dang_ngay_hien_thi(val):
    digits = _lay_12_so(val)
    if len(digits) != 12:
        return _chuan_hoa_gia_tri(val)
    return f"{digits[6:8]}/{digits[4:6]}/{digits[0:4]} {digits[8:10]}:{digits[10:12]}"


def _doc_phut_dich_vu(phut_theo_dich_vu):
    ket_qua = {}
    phut_theo_dich_vu = phut_theo_dich_vu or {}
    for ten, phut_mac_dinh, mau_ten in CAU_HINH_DICH_VU_MAC_DINH:
        raw = phut_theo_dich_vu.get(ten, phut_mac_dinh)
        try:
            phut = int(float(str(raw).strip().replace(",", ".")))
        except Exception:
            phut = phut_mac_dinh
        if phut <= 0:
            phut = phut_mac_dinh
        ket_qua[ten] = {"phut": phut, "mau_ten": mau_ten}
    return ket_qua


def _tim_dich_vu_cau_hinh(ten_dich_vu, cau_hinh):
    ten_chuan = _chuan_hoa_text(ten_dich_vu)
    for ten_nhom, thong_tin in cau_hinh.items():
        for mau in thong_tin["mau_ten"]:
            if _chuan_hoa_text(mau) in ten_chuan:
                return ten_nhom
    return ""


def chuan_doan_xml3_so_luot_yhct(file_xml3, phut_theo_dich_vu=None, callback_status=None):
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
        callback_status("Đang tính số lượt DVKT theo người thực hiện, dịch vụ và ngày...", "blue")

    cau_hinh = _doc_phut_dich_vu(phut_theo_dich_vu)
    df_work = df[COT_BAT_BUOC].copy()
    for cot in COT_BAT_BUOC:
        df_work[cot] = df_work[cot].apply(_chuan_hoa_gia_tri)

    df_work["_NHOM_DVKT"] = df_work["TEN_DICH_VU"].apply(lambda val: _tim_dich_vu_cau_hinh(val, cau_hinh))
    df_work["_NGAY_KEY"] = df_work.apply(_lay_ngay_key, axis=1)
    df_work["_NGUOI_KEY"] = df_work["NGUOI_THUC_HIEN"].apply(_chuan_hoa_text)
    df_work["_TEN_DICH_VU_KEY"] = df_work["TEN_DICH_VU"].apply(_chuan_hoa_text)

    ma_nhom = df_work["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    mask_xet = (ma_nhom == "18") & (df_work["_NHOM_DVKT"] != "") & (df_work["_NGUOI_KEY"] != "") & (df_work["_NGAY_KEY"] != "")
    so_luot = (
        df_work.loc[mask_xet]
        .groupby(["_NGUOI_KEY", "_NHOM_DVKT", "_NGAY_KEY"])["TEN_DICH_VU"]
        .transform("size")
    )
    df_work["_SO_LUOT"] = 0
    df_work.loc[mask_xet, "_SO_LUOT"] = so_luot.to_numpy()

    df_work["_GIOI_HAN"] = df_work["_NHOM_DVKT"].map(
        lambda ten: math.floor(480 / cau_hinh[ten]["phut"]) if ten in cau_hinh else 0
    )
    mask_loi = mask_xet & (df_work["_SO_LUOT"] > df_work["_GIOI_HAN"])

    df_work["LOI"] = ""
    df_work.loc[mask_loi, "LOI"] = LOI_VUOT_GIOI_HAN
    for cot_ngay in ["NGAY_KQ", "NGAY_TH_YL"]:
        df_work[cot_ngay] = df_work[cot_ngay].apply(_dinh_dang_ngay_hien_thi)

    mask_hien_thi = mask_xet
    df_work = df_work.loc[mask_hien_thi].copy()
    df_work["_UU_TIEN_LOI"] = df_work["LOI"].ne("").astype(int)
    df_ket_qua = (
        df_work.sort_values(["_UU_TIEN_LOI", "NGUOI_THUC_HIEN", "TEN_DICH_VU"], ascending=[False, True, True], kind="stable")
        .drop(columns=["MA_NHOM", "_NHOM_DVKT", "_NGAY_KEY", "_NGUOI_KEY", "_TEN_DICH_VU_KEY", "_SO_LUOT", "_GIOI_HAN", "_UU_TIEN_LOI"])
        .loc[:, COT_HIEN_THI]
        .reset_index(drop=True)
    )

    if callback_status:
        so_loi = int(df_ket_qua["LOI"].eq(LOI_VUOT_GIOI_HAN).sum())
        callback_status(f"Chuẩn đoán xong: {len(df_ket_qua)} dòng, {so_loi} lỗi.", "green")

    return df_ket_qua


def xuat_excel_chuan_doan_xml3_so_luot_yhct(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="SoLuot_YHCT", index=False)
        ws = writer.sheets["SoLuot_YHCT"]
        fill_loi = PatternFill(fill_type="solid", fgColor="FFF59D")

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"NGUOI_THUC_HIEN", "TEN_DICH_VU", "LOI"}:
                width = 36 if ten_cot != "LOI" else 48
            elif ten_cot in {"NGAY_KQ", "NGAY_TH_YL"}:
                width = 18
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

        headers = [cell.value for cell in ws[1]]
        loi_col_idx = headers.index("LOI") + 1 if "LOI" in headers else None
        if loi_col_idx:
            for row in ws.iter_rows(min_row=2):
                if row[loi_col_idx - 1].value == LOI_VUOT_GIOI_HAN:
                    for cell in row:
                        cell.fill = fill_loi

    return file_xuat
