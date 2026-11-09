import os

import pandas as pd
from openpyxl.styles import Font, PatternFill

from core.utils import doc_file_thong_minh, lam_sach_chuoi


HEADER_FILL = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
MA_NHOM_LOAI_BO = {1, 3, 8, 10, 13, 14, 15}

COT_HIEN_THI = [
    "MA_LK",
    "MA_BA",
    "HO_TEN",
    "NGAY_YL",
    "NGAY_KQ",
    "MA_DICH_VU",
    "TEN_DICH_VU",
    "SOPHIEU",
]


def log(callback_status, message, color="blue"):
    print(message)
    if callback_status:
        callback_status(message, color)


def tao_file_khong_trung(duong_dan):
    goc, ext = os.path.splitext(duong_dan)
    dem = 1
    ket_qua = duong_dan
    while os.path.exists(ket_qua):
        ket_qua = f"{goc} ({dem}){ext}"
        dem += 1
    return ket_qua


def kiem_tra_cot_bat_buoc(df, cot_bat_buoc, ten_file):
    cot_thieu = [cot for cot in cot_bat_buoc if cot not in df.columns]
    if cot_thieu:
        raise ValueError(f"{ten_file} thiếu các cột bắt buộc: {', '.join(cot_thieu)}")


def chuan_hoa_chuoi(df, cot):
    if cot not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    return df[cot].map(lam_sach_chuoi)


def chuan_hoa_ma_nhom(val):
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def dinh_dang_thoi_gian_hien_thi(series_raw):
    dt = pd.to_datetime(series_raw, format="%Y%m%d%H%M", errors="coerce")
    return dt.dt.strftime("%d/%m/%Y %H:%M").where(dt.notna(), series_raw)


def xuat_excel_ket_qua(df_ket_qua, file_xuat):
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="KetQua_CD72", index=False)
        worksheet = writer.sheets["KetQua_CD72"]

        for cell in worksheet[1]:
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)

        do_rong = {
            "MA_LK": 16,
            "MA_BA": 16,
            "HO_TEN": 28,
            "NGAY_YL": 18,
            "NGAY_KQ": 18,
            "MA_DICH_VU": 24,
            "TEN_DICH_VU": 42,
            "SOPHIEU": 18,
        }
        for col in worksheet.columns:
            ten_cot = col[0].value
            worksheet.column_dimensions[col[0].column_letter].width = do_rong.get(ten_cot, 18)


def xu_ly_chuyen_de_72(file_xml3, file_xuat, callback_status):
    try:
        file_xuat = tao_file_khong_trung(file_xuat)

        log(callback_status, "[1/4] Đang đọc dữ liệu XML3...")
        df_xml3 = doc_file_thong_minh(file_xml3)
        df_xml3.columns = df_xml3.columns.str.strip()

        kiem_tra_cot_bat_buoc(
            df_xml3,
            [
                "MA_LK",
                "MA_BA",
                "HO_TEN",
                "MA_THE",
                "MA_NHOM",
                "NGAY_YL",
                "NGAY_TH_YL",
                "NGAY_KQ",
                "MA_DICH_VU",
                "TEN_DICH_VU",
                "SOPHIEU",
            ],
            "XML3",
        )

        log(callback_status, "[2/4] Chuẩn hóa dữ liệu và loại bỏ MA_NHOM không xét...")
        df3 = df_xml3.copy()
        for cot in [
            "MA_LK",
            "MA_BA",
            "HO_TEN",
            "MA_THE",
            "MA_NHOM",
            "NGAY_YL",
            "NGAY_TH_YL",
            "NGAY_KQ",
            "MA_DICH_VU",
            "TEN_DICH_VU",
            "SOPHIEU",
        ]:
            df3[cot] = chuan_hoa_chuoi(df3, cot)

        ma_nhom_series = df3["MA_NHOM"].apply(chuan_hoa_ma_nhom)
        df3 = df3.loc[~ma_nhom_series.isin(MA_NHOM_LOAI_BO)].copy()

        # Tránh bắt nhầm khi thiếu thông tin khóa phát hiện trùng.
        for cot in ["MA_BA", "HO_TEN", "MA_THE", "NGAY_TH_YL", "NGAY_KQ", "MA_DICH_VU"]:
            df3 = df3[df3[cot] != ""]

        log(callback_status, "[3/4] Phát hiện trùng theo MA_BA, HO_TEN, MA_THE, NGAY_TH_YL, NGAY_KQ...")
        khoa_nhom = ["MA_BA", "HO_TEN", "MA_THE", "NGAY_TH_YL", "NGAY_KQ"]
        thong_ke_nhom = (
            df3.groupby(khoa_nhom, dropna=False)
            .agg(SO_DONG=("MA_DICH_VU", "size"), SO_MA_DV_KHAC_NHAU=("MA_DICH_VU", "nunique"))
            .reset_index()
        )
        khoa_loi = thong_ke_nhom[
            (thong_ke_nhom["SO_DONG"] >= 2) & (thong_ke_nhom["SO_MA_DV_KHAC_NHAU"] >= 2)
        ][khoa_nhom]

        df_loi = df3.merge(khoa_loi, on=khoa_nhom, how="inner")
        df_loi["NGAY_YL"] = dinh_dang_thoi_gian_hien_thi(df_loi["NGAY_YL"])
        df_loi["NGAY_KQ"] = dinh_dang_thoi_gian_hien_thi(df_loi["NGAY_KQ"])

        if df_loi.empty:
            df_ket_qua = pd.DataFrame(columns=COT_HIEN_THI)
        else:
            df_ket_qua = df_loi[COT_HIEN_THI].sort_values(
                by=["MA_BA", "HO_TEN", "NGAY_KQ", "MA_DICH_VU"],
                na_position="last",
            )
            df_ket_qua = df_ket_qua.reset_index(drop=True)

        log(callback_status, f"[4/4] Xuất file kết quả: {len(df_ket_qua):,} dòng.")
        xuat_excel_ket_qua(df_ket_qua, file_xuat)

        log(callback_status, "Hoàn thành kiểm tra Chuyên đề 72.", "green")
        return True, file_xuat
    except Exception as e:
        log(callback_status, "Lỗi khi xử lý Chuyên đề 72.", "red")
        return False, str(e)
