import os

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill

from core.utils import doc_file_thong_minh, lam_sach_chuoi


RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
HEADER_FILL = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")


def chuyen_ngay_xml(value):
    value_str = lam_sach_chuoi(value)
    if not value_str:
        return pd.NaT

    for fmt in ("%Y%m%d%H%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y%m%d"):
        ngay = pd.to_datetime(value_str, format=fmt, errors="coerce")
        if pd.notna(ngay):
            return ngay

    return pd.to_datetime(value_str, errors="coerce")


def dinh_dang_ngay(value):
    if pd.isna(value):
        return ""
    return value.strftime("%d/%m/%Y %H:%M")


def cat_ngan_noi_dung(value, max_len=120):
    text = lam_sach_chuoi(value).replace("\r", " ").replace("\n", " ")
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def lay_ten_dich_vu_xml3(row):
    ten_dich_vu = lam_sach_chuoi(row.get("TEN_DICH_VU", ""))
    if ten_dich_vu:
        return ten_dich_vu
    return lam_sach_chuoi(row.get("TEN_VAT_TU", ""))


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
        raise ValueError(f"{ten_file} thieu cac cot bat buoc: {', '.join(cot_thieu)}")


def tao_bang_loi_xml3(df_xml3_merged):
    df_loi = df_xml3_merged[
        df_xml3_merged["XML3_NGAY_KQ_DT"].notna() &
        df_xml3_merged["XML1_NGAY_RA_DT"].notna() &
        (df_xml3_merged["XML3_NGAY_KQ_DT"] > df_xml3_merged["XML1_NGAY_RA_DT"])
    ].copy()
    if df_loi.empty:
        return pd.DataFrame()

    return pd.DataFrame(
        {
            "MA_LK": df_loi["MA_LK"].map(lam_sach_chuoi),
            "HO_TEN": df_loi["HO_TEN"].map(lam_sach_chuoi),
            "NGAY_RA_CHUAN": df_loi["XML1_NGAY_RA_DT"].map(dinh_dang_ngay),
            "TEN_DICH_VU_LOI": df_loi.apply(lay_ten_dich_vu_xml3, axis=1),
            "NGAY_KQ_LOI": df_loi["XML3_NGAY_KQ_DT"].map(dinh_dang_ngay),
            "PHAN_LOAI_LOI": "Lỗi CĐ69: Thời gian trả kết quả dịch vụ muộn",
            "FILE_NGUON": "XML3",
            "ERROR_COLUMN": "NGAY_KQ_LOI",
        }
    )


def tao_bang_loi_xml4(df_xml4_merged):
    df_loi = df_xml4_merged[
        df_xml4_merged["XML4_NGAY_KQ_DT"].notna() &
        df_xml4_merged["XML1_NGAY_RA_DT"].notna() &
        (df_xml4_merged["XML4_NGAY_KQ_DT"] > df_xml4_merged["XML1_NGAY_RA_DT"])
    ].copy()
    if df_loi.empty:
        return pd.DataFrame()

    return pd.DataFrame(
        {
            "MA_LK": df_loi["MA_LK"].map(lam_sach_chuoi),
            "HO_TEN": df_loi["HO_TEN"].map(lam_sach_chuoi),
            "NGAY_RA_CHUAN": df_loi["XML1_NGAY_RA_DT"].map(dinh_dang_ngay),
            "TEN_DICH_VU_LOI": df_loi["TEN_CHI_SO"].map(lam_sach_chuoi),
            "NGAY_KQ_LOI": df_loi["XML4_NGAY_KQ_DT"].map(dinh_dang_ngay),
            "PHAN_LOAI_LOI": "Lỗi CĐ69: Thời gian trả kết quả dịch vụ muộn",
            "FILE_NGUON": "XML4",
            "ERROR_COLUMN": "NGAY_KQ_LOI",
        }
    )


def xu_ly_chuyen_de_69(file_xml1, file_xml3, file_xml4, file_xml5, file_xml7, file_xuat, tu_ngay, den_ngay, callback_status):
    try:
        file_xuat = tao_file_khong_trung(file_xuat)

        callback_status("Dang doc du lieu XML1, XML3, XML4...", "blue")
        df_xml1 = doc_file_thong_minh(file_xml1)
        df_xml3 = doc_file_thong_minh(file_xml3)
        df_xml4 = doc_file_thong_minh(file_xml4)

        df_xml1.columns = df_xml1.columns.str.strip()
        df_xml3.columns = df_xml3.columns.str.strip()
        df_xml4.columns = df_xml4.columns.str.strip()

        kiem_tra_cot_bat_buoc(df_xml1, ["MA_LK", "HO_TEN", "NGAY_RA"], "XML1")
        kiem_tra_cot_bat_buoc(df_xml3, ["MA_LK", "TEN_DICH_VU", "TEN_VAT_TU", "NGAY_KQ"], "XML3")
        kiem_tra_cot_bat_buoc(df_xml4, ["MA_LK", "TEN_CHI_SO", "NGAY_KQ"], "XML4")

        callback_status("Dang chuan hoa thoi gian va loc theo ngay ra vien...", "blue")
        df_xml1 = df_xml1.copy()
        df_xml1["MA_LK"] = df_xml1["MA_LK"].apply(lam_sach_chuoi)
        df_xml1["HO_TEN"] = df_xml1["HO_TEN"].apply(lam_sach_chuoi)
        df_xml1["XML1_NGAY_RA_DT"] = df_xml1["NGAY_RA"].apply(chuyen_ngay_xml)
        df_xml1 = df_xml1[df_xml1["MA_LK"] != ""].copy()

        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        df_xml1 = df_xml1[
            (df_xml1["XML1_NGAY_RA_DT"] >= tu_ngay_dt) & (df_xml1["XML1_NGAY_RA_DT"] <= den_ngay_dt)
        ].copy()
        df_xml1 = df_xml1[["MA_LK", "HO_TEN", "XML1_NGAY_RA_DT"]].drop_duplicates(subset=["MA_LK"])

        df_xml3 = df_xml3.copy()
        df_xml3["MA_LK"] = df_xml3["MA_LK"].apply(lam_sach_chuoi)
        df_xml3["XML3_NGAY_KQ_DT"] = df_xml3["NGAY_KQ"].apply(chuyen_ngay_xml)

        df_xml4 = df_xml4.copy()
        df_xml4["MA_LK"] = df_xml4["MA_LK"].apply(lam_sach_chuoi)
        df_xml4["XML4_NGAY_KQ_DT"] = df_xml4["NGAY_KQ"].apply(chuyen_ngay_xml)

        callback_status("Dang doi chieu XML3 va XML4 voi ngay ra vien cua XML1...", "blue")
        xml3_merged = df_xml3.merge(df_xml1, on="MA_LK", how="inner", suffixes=("", "_XML1"))
        xml4_merged = df_xml4.merge(df_xml1, on="MA_LK", how="inner", suffixes=("", "_XML1"))
        if "HO_TEN_XML1" in xml3_merged.columns:
            xml3_merged["HO_TEN"] = xml3_merged["HO_TEN_XML1"]
        if "HO_TEN_XML1" in xml4_merged.columns:
            xml4_merged["HO_TEN"] = xml4_merged["HO_TEN_XML1"]

        callback_status("Dang phan loai cac dong loi Chuyen de 69...", "blue")
        df_ket_qua = pd.concat([tao_bang_loi_xml3(xml3_merged), tao_bang_loi_xml4(xml4_merged)], ignore_index=True)
        if df_ket_qua.empty:
            df_ket_qua = pd.DataFrame(
                columns=[
                    "MA_LK",
                    "HO_TEN",
                    "NGAY_RA_CHUAN",
                    "TEN_DICH_VU_LOI",
                    "NGAY_KQ_LOI",
                    "PHAN_LOAI_LOI",
                    "FILE_NGUON",
                    "ERROR_COLUMN",
                ]
            )

        cot_ket_qua = [
            "MA_LK",
            "HO_TEN",
            "NGAY_RA_CHUAN",
            "TEN_DICH_VU_LOI",
            "NGAY_KQ_LOI",
            "PHAN_LOAI_LOI",
            "FILE_NGUON",
            "ERROR_COLUMN",
        ]
        df_ket_qua = df_ket_qua.reindex(columns=cot_ket_qua)

        if not df_ket_qua.empty:
            df_ket_qua = df_ket_qua.sort_values(
                by=["FILE_NGUON", "MA_LK", "NGAY_KQ_LOI"],
                ascending=[True, True, True],
            ).reset_index(drop=True)
        df_ket_qua.insert(0, "STT", range(1, len(df_ket_qua) + 1))

        callback_status("Dang xuat file Excel tong hop...", "blue")
        with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
            df_ket_qua.to_excel(writer, sheet_name="KetQua_CD69", index=False)
            worksheet = writer.sheets["KetQua_CD69"]

            for cell in worksheet[1]:
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)

            cot_theo_tieu_de = {worksheet.cell(row=1, column=i).value: i for i in range(1, worksheet.max_column + 1)}
            cot_ngay_kq_loi = cot_theo_tieu_de.get("NGAY_KQ_LOI")

            for dong_excel in range(2, worksheet.max_row + 1):
                for cot in range(1, worksheet.max_column + 1):
                    worksheet.cell(row=dong_excel, column=cot).fill = RED_FILL

                if cot_ngay_kq_loi:
                    worksheet.cell(row=dong_excel, column=cot_ngay_kq_loi).font = Font(color="9C0006", bold=True)

            do_rong = {
                "STT": 8,
                "MA_LK": 20,
                "HO_TEN": 28,
                "NGAY_RA_CHUAN": 18,
                "TEN_DICH_VU_LOI": 42,
                "NGAY_KQ_LOI": 18,
                "PHAN_LOAI_LOI": 34,
                "FILE_NGUON": 12,
                "ERROR_COLUMN": 16,
            }

            for col in worksheet.columns:
                ten_cot = col[0].value
                worksheet.column_dimensions[col[0].column_letter].width = do_rong.get(ten_cot, 18)

        callback_status("Hoan thanh kiem tra Chuyen de 69.", "green")
        try:
            os.startfile(file_xuat)
        except Exception:
            pass
        return True, file_xuat
    except Exception as e:
        callback_status("Loi khi xu ly Chuyen de 69.", "red")
        return False, str(e)
