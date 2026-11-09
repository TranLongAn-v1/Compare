import os
import re
import unicodedata
import warnings

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh, lam_sach_chuoi

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)


def safe_float(val):
    try:
        if pd.isna(val) or val == "":
            return 0.0
        clean_val = str(val).replace(",", "").strip()
        return float(clean_val)
    except Exception:
        return 0.0


def chuan_hoa_ten_cot(val):
    raw = lam_sach_chuoi(val).lower()
    raw = raw.replace("đ", "d").replace("Ä'", "d")
    no_d = "".join(ch for ch in unicodedata.normalize("NFD", raw) if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]", "", no_d)


def tim_cot(df, *ten_uu_tien):
    for ten in ten_uu_tien:
        if ten in df.columns:
            return ten
    cot_chuan = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}
    for ten in ten_uu_tien:
        ten_chuan = re.sub(r"\s+", " ", ten.strip().lower())
        if ten_chuan in cot_chuan:
            return cot_chuan[ten_chuan]
    cot_khong_dau = {chuan_hoa_ten_cot(c): c for c in df.columns}
    for ten in ten_uu_tien:
        ten_khong_dau = chuan_hoa_ten_cot(ten)
        if ten_khong_dau in cot_khong_dau:
            return cot_khong_dau[ten_khong_dau]
    raise KeyError(f"Khong tim thay cot nao trong danh sach: {ten_uu_tien}")


def doc_file_mau01_thong_minh(file_mau01):
    raw = pd.read_excel(file_mau01, engine="openpyxl", header=None)
    nhom_bat_buoc = [
        {"mathe", "mathebhyt"},
        {"hovaten", "hoten"},
        {"ngayvao"},
        {"ngayra"},
        {"tongchibv", "ttongchibv", "tongchi"},
    ]

    header_idx = None
    diem_tot_nhat = 0
    for idx, row in raw.head(30).iterrows():
        ten_cot_tren_dong = {chuan_hoa_ten_cot(v) for v in row.tolist()}
        diem = sum(1 for nhom in nhom_bat_buoc if ten_cot_tren_dong & nhom)
        if diem > diem_tot_nhat:
            diem_tot_nhat = diem
            header_idx = idx

    if header_idx is None or diem_tot_nhat < 4:
        return doc_file_thong_minh(file_mau01)

    columns = []
    for i, col in enumerate(raw.iloc[header_idx].tolist()):
        ten = lam_sach_chuoi(col)
        columns.append(ten if ten else f"__EMPTY_{i}")

    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = columns
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return df


def parse_datetime_flexible(val):
    s = lam_sach_chuoi(val)
    if not s:
        return pd.NaT
    # Mau 01 co truong hop gio 24h nhung van kem hau to SA/CH, can bo qua hau to nay.
    s = re.sub(r"\s+(SA|CH)\s*$", "", s, flags=re.IGNORECASE)
    dt = pd.to_datetime(s, format="%d/%m/%Y %H:%M", errors="coerce")
    if pd.notna(dt):
        return dt
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return dt
    return pd.NaT


def dinh_dang_dt_hien_thi(dt_val):
    if pd.isna(dt_val):
        return ""
    return pd.Timestamp(dt_val).strftime("%d/%m/%Y %H:%M")


def chuan_hoa_ma_the(val):
    raw = lam_sach_chuoi(val).upper()
    return re.sub(r"[^A-Z0-9]", "", raw)


def chuan_hoa_ho_ten(val):
    raw = re.sub(r"\s+", " ", lam_sach_chuoi(val).upper()).strip()
    if not raw:
        return ""
    no_d = raw.replace("Đ", "D")
    no_d = "".join(ch for ch in unicodedata.normalize("NFD", no_d) if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^A-Z0-9 ]", "", no_d)


def _so_sanh_datetime(a, b):
    if pd.isna(a) and pd.isna(b):
        return 0
    if pd.isna(a) or pd.isna(b):
        return 10**9
    return abs(int((pd.Timestamp(a) - pd.Timestamp(b)).total_seconds()))


def tao_du_lieu_doi_chieu_giam_dinh_xml_mau01(file_giam_dinh_xml, file_mau01, tu_ngay, den_ngay, callback_status):
    try:
        callback_status("Dang doc du lieu tu Excel...", "blue")
        df_gdx = doc_file_thong_minh(file_giam_dinh_xml)
        df_mau01 = doc_file_mau01_thong_minh(file_mau01)

        cgdx_ma_the = tim_cot(df_gdx, "Mã thẻ", "Ma the", "MA_THE", "MATHE")
        cgdx_ho_ten = tim_cot(df_gdx, "Họ tên", "Ho ten", "Họ và tên", "Ho va ten", "HO_TEN", "HOTEN")
        cgdx_ngay_vao = tim_cot(df_gdx, "Ngày vào", "Ngay vao", "NGAY_VAO", "NGAYVAO")
        cgdx_ngay_ra = tim_cot(df_gdx, "Ngày ra", "Ngay ra", "NGAY_RA", "NGAYRA")
        cgdx_tong_chi = tim_cot(df_gdx, "Tổng chi", "Tong chi", "T_TONGCHI", "TONG_CHI", "TONGCHI")
        cgdx_bhtt = tim_cot(df_gdx, "Bảo hiểm TT", "Bao hiem TT", "T_BHTT", "BAOHIEMTT", "TBHTT")

        cm01_ma_the = tim_cot(df_mau01, "Mã thẻ", "Ma the", "MA_THE", "Mã thẻ BHYT", "Ma the BHYT", "MATHE")
        cm01_ho_ten = tim_cot(df_mau01, "Họ và tên", "Ho va ten", "HO_TEN", "Họ tên", "Ho ten", "HOTEN")
        cm01_ngay_vao = tim_cot(df_mau01, "Ngày vào", "Ngay vao", "NGAY_VAO", "NGAYVAO")
        cm01_ngay_ra = tim_cot(df_mau01, "Ngày ra", "Ngay ra", "NGAY_RA", "NGAYRA")
        cm01_tong_chi = tim_cot(
            df_mau01,
            "Tổng chi BV",
            "Tong chi BV",
            "TONG_CHI_BV",
            "T_TONGCHI_BV",
            "Tổng chi",
            "Tong chi",
            "TONGCHIBV",
            "TTONGCHIBV",
        )
        cm01_tong_chi_bh = tim_cot(
            df_mau01,
            "Tổng chi BH",
            "Tong chi BH",
            "TONG_CHI_BH",
            "T_TONGCHI_BH",
            "TONGCHIBH",
            "TTONGCHIBH",
        )
        cm01_bhtt = tim_cot(df_mau01, "Bảo hiểm TT", "Bao hiem TT", "T_BHTT", "BAOHIEMTT", "TBHTT")

        df_gdx = df_gdx.rename(
            columns={
                cgdx_ma_the: "MA_THE_GDX",
                cgdx_ho_ten: "HO_TEN_GDX",
                cgdx_ngay_vao: "NGAY_VAO_GDX",
                cgdx_ngay_ra: "NGAY_RA_GDX",
                cgdx_tong_chi: "TONG_CHI_GDX",
                cgdx_bhtt: "T_BHTT_GDX",
            }
        )
        df_mau01 = df_mau01.rename(
            columns={
                cm01_ma_the: "MA_THE_M01",
                cm01_ho_ten: "HO_TEN_M01",
                cm01_ngay_vao: "NGAY_VAO_M01",
                cm01_ngay_ra: "NGAY_RA_M01",
                cm01_tong_chi: "TONG_CHI_M01",
                cm01_tong_chi_bh: "TONG_CHI_BH_M01",
                cm01_bhtt: "T_BHTT_M01",
            }
        )

        callback_status("Dang chuan hoa du lieu...", "blue")
        for col in ["MA_THE_GDX", "HO_TEN_GDX"]:
            df_gdx[col] = df_gdx[col].apply(lam_sach_chuoi)
        for col in ["MA_THE_M01", "HO_TEN_M01"]:
            df_mau01[col] = df_mau01[col].apply(lam_sach_chuoi)

        df_gdx["MA_THE_CHUAN"] = df_gdx["MA_THE_GDX"].apply(chuan_hoa_ma_the)
        df_mau01["MA_THE_CHUAN"] = df_mau01["MA_THE_M01"].apply(chuan_hoa_ma_the)
        df_gdx["HO_TEN_CHUAN"] = df_gdx["HO_TEN_GDX"].apply(chuan_hoa_ho_ten)
        df_mau01["HO_TEN_CHUAN"] = df_mau01["HO_TEN_M01"].apply(chuan_hoa_ho_ten)

        df_gdx["Ngay_Vao_DT"] = df_gdx["NGAY_VAO_GDX"].apply(parse_datetime_flexible)
        df_gdx["Ngay_Ra_DT"] = df_gdx["NGAY_RA_GDX"].apply(parse_datetime_flexible)
        df_gdx["NGAY_VAO_CHUAN"] = df_gdx["Ngay_Vao_DT"].apply(dinh_dang_dt_hien_thi)
        df_gdx["NGAY_RA_CHUAN"] = df_gdx["Ngay_Ra_DT"].apply(dinh_dang_dt_hien_thi)

        df_mau01["Ngay_Vao_DT"] = df_mau01["NGAY_VAO_M01"].apply(parse_datetime_flexible)
        df_mau01["Ngay_Ra_DT"] = df_mau01["NGAY_RA_M01"].apply(parse_datetime_flexible)
        df_mau01["NGAY_VAO_CHUAN"] = df_mau01["Ngay_Vao_DT"].apply(dinh_dang_dt_hien_thi)
        df_mau01["NGAY_RA_CHUAN"] = df_mau01["Ngay_Ra_DT"].apply(dinh_dang_dt_hien_thi)

        df_gdx_valid = df_gdx[(df_gdx["HO_TEN_CHUAN"] != "")].copy()
        df_mau01_valid = df_mau01[(df_mau01["HO_TEN_CHUAN"] != "")].copy()

        callback_status("Dang loc theo khoang thoi gian...", "blue")
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        df_gdx_valid = df_gdx_valid.loc[
            (df_gdx_valid["Ngay_Ra_DT"] >= tu_ngay_dt) & (df_gdx_valid["Ngay_Ra_DT"] <= den_ngay_dt)
        ].copy()
        df_mau01_valid = df_mau01_valid.loc[
            (df_mau01_valid["Ngay_Ra_DT"] >= tu_ngay_dt) & (df_mau01_valid["Ngay_Ra_DT"] <= den_ngay_dt)
        ].copy()

        callback_status("Dang doi chieu du lieu...", "blue")
        rows = []
        ds_gdx = df_gdx_valid.reset_index(drop=True).to_dict("records")
        ds_m01 = df_mau01_valid.reset_index(drop=True).to_dict("records")
        da_ghep_gdx = set()

        for row_m01 in ds_m01:
            idx_tot_nhat = None
            diem_tot_nhat = None

            for idx, row_gdx in enumerate(ds_gdx):
                if idx in da_ghep_gdx:
                    continue
                if row_gdx.get("HO_TEN_CHUAN", "") != row_m01.get("HO_TEN_CHUAN", ""):
                    continue

                cung_ma_the = row_gdx.get("MA_THE_CHUAN", "") == row_m01.get("MA_THE_CHUAN", "")
                cung_ngay = (
                    dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Vao_DT", pd.NaT)) == dinh_dang_dt_hien_thi(row_m01.get("Ngay_Vao_DT", pd.NaT))
                    and dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Ra_DT", pd.NaT)) == dinh_dang_dt_hien_thi(row_m01.get("Ngay_Ra_DT", pd.NaT))
                )

                if cung_ma_the:
                    muc_uu_tien = 1
                elif cung_ngay:
                    muc_uu_tien = 2
                else:
                    muc_uu_tien = 3

                khoang_cach_ngay = _so_sanh_datetime(row_gdx.get("Ngay_Vao_DT", pd.NaT), row_m01.get("Ngay_Vao_DT", pd.NaT))
                diem = (muc_uu_tien, khoang_cach_ngay)
                if diem_tot_nhat is None or diem < diem_tot_nhat:
                    diem_tot_nhat = diem
                    idx_tot_nhat = idx

            if idx_tot_nhat is not None:
                da_ghep_gdx.add(idx_tot_nhat)
                row_gdx = ds_gdx[idx_tot_nhat]
                ton_tai_gdx = True
                ton_tai_m01 = True
            else:
                row_gdx = {}
                ton_tai_gdx = False
                ton_tai_m01 = True

            ngay_vao_gdx = dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Vao_DT", pd.NaT))
            ngay_ra_gdx = dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Ra_DT", pd.NaT))
            ngay_vao_m01 = dinh_dang_dt_hien_thi(row_m01.get("Ngay_Vao_DT", pd.NaT))
            ngay_ra_m01 = dinh_dang_dt_hien_thi(row_m01.get("Ngay_Ra_DT", pd.NaT))

            tong_gdx = safe_float(row_gdx.get("TONG_CHI_GDX", 0))
            tong_m01 = safe_float(row_m01.get("TONG_CHI_M01", 0))
            tong_bh_m01 = safe_float(row_m01.get("TONG_CHI_BH_M01", 0))
            bhtt_gdx = safe_float(row_gdx.get("T_BHTT_GDX", 0))
            bhtt_m01 = safe_float(row_m01.get("T_BHTT_M01", 0))
            ma_the_gdx = lam_sach_chuoi(row_gdx.get("MA_THE_GDX", ""))
            ma_the_m01 = lam_sach_chuoi(row_m01.get("MA_THE_M01", ""))
            ma_the_gdx_chuan = lam_sach_chuoi(row_gdx.get("MA_THE_CHUAN", ""))
            ma_the_m01_chuan = lam_sach_chuoi(row_m01.get("MA_THE_CHUAN", ""))

            lech_tien = []
            if abs(tong_gdx - tong_m01) > 1.0:
                lech_tien.append("Tong chi")
            if abs(bhtt_gdx - bhtt_m01) > 1.0:
                lech_tien.append("T_BHTT")

            phan_loai, mau_to = "Trung khop", "none"
            if ton_tai_m01 and not ton_tai_gdx:
                phan_loai, mau_to = "Co tren Mau 01 nhung khong co tren Giam dinh XML", "blue"
            elif ma_the_gdx_chuan != ma_the_m01_chuan:
                phan_loai, mau_to = "Lech ma the", "purple"
            elif (ngay_vao_gdx != ngay_vao_m01) or (ngay_ra_gdx != ngay_ra_m01):
                phan_loai, mau_to = "Lech ngay vao/ngay ra", "orange"
            elif lech_tien:
                phan_loai, mau_to = f"Lech tien: {', '.join(lech_tien)}", "red"

            rows.append(
                {
                    "HO_TEN": row_m01.get("HO_TEN_M01", "") if ton_tai_m01 else row_gdx.get("HO_TEN_GDX", ""),
                    "MA_THE (Giam dinh XML)": ma_the_gdx,
                    "MA_THE (Mau 01)": ma_the_m01,
                    "NGAY_VAO (Giam dinh XML)": ngay_vao_gdx,
                    "NGAY_VAO (Mau 01)": ngay_vao_m01,
                    "NGAY_RA (Giam dinh XML)": ngay_ra_gdx,
                    "NGAY_RA (Mau 01)": ngay_ra_m01,
                    "TONG_CHI (Giam dinh XML)": tong_gdx,
                    "TONG_CHI_BV (Mau 01)": tong_m01,
                    "TONG_CHI_BH (Mau 01)": tong_bh_m01,
                    "T_BHTT (Giam dinh XML)": bhtt_gdx,
                    "T_BHTT (Mau 01)": bhtt_m01,
                    "PHAN_LOAI": phan_loai,
                    "color": mau_to,
                }
            )

        for idx, row_gdx in enumerate(ds_gdx):
            if idx in da_ghep_gdx:
                continue
            rows.append(
                {
                    "HO_TEN": row_gdx.get("HO_TEN_GDX", ""),
                    "MA_THE (Giam dinh XML)": lam_sach_chuoi(row_gdx.get("MA_THE_GDX", "")),
                    "MA_THE (Mau 01)": "",
                    "NGAY_VAO (Giam dinh XML)": dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Vao_DT", pd.NaT)),
                    "NGAY_VAO (Mau 01)": "",
                    "NGAY_RA (Giam dinh XML)": dinh_dang_dt_hien_thi(row_gdx.get("Ngay_Ra_DT", pd.NaT)),
                    "NGAY_RA (Mau 01)": "",
                    "TONG_CHI (Giam dinh XML)": safe_float(row_gdx.get("TONG_CHI_GDX", 0)),
                    "TONG_CHI_BV (Mau 01)": 0.0,
                    "TONG_CHI_BH (Mau 01)": 0.0,
                    "T_BHTT (Giam dinh XML)": safe_float(row_gdx.get("T_BHTT_GDX", 0)),
                    "T_BHTT (Mau 01)": 0.0,
                    "PHAN_LOAI": "Co tren Giam dinh XML nhung khong co tren Mau 01",
                    "color": "yellow",
                }
            )

        df_tong_hop = pd.DataFrame(rows)
        if not df_tong_hop.empty:
            thu_tu = {
                "Co tren Mau 01 nhung khong co tren Giam dinh XML": 1,
                "Co tren Giam dinh XML nhung khong co tren Mau 01": 2,
                "Lech ma the": 3,
                "Lech ngay vao/ngay ra": 4,
                "Lech tien": 5,
                "Trung khop": 6,
            }
            df_tong_hop["_THU_TU"] = df_tong_hop["PHAN_LOAI"].apply(
                lambda val: thu_tu["Lech tien"] if str(val).startswith("Lech tien") else thu_tu.get(str(val), 99)
            )
            df_tong_hop = df_tong_hop.sort_values(["_THU_TU", "PHAN_LOAI"]).drop(columns=["_THU_TU"]).reset_index(drop=True)
            df_tong_hop.insert(0, "STT", range(1, len(df_tong_hop) + 1))

        callback_status("Hoan thanh doi chieu.", "green")
        return True, df_tong_hop
    except Exception as e:
        callback_status("Loi xay ra.", "red")
        return False, str(e)


def xuat_excel_tong_hop(df_tong_hop, file_xuat):
    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    df_xuat_excel = df_tong_hop.drop(columns=["color"], errors="ignore")
    cot_tien = [
        "TONG_CHI (Giam dinh XML)",
        "TONG_CHI_BV (Mau 01)",
        "TONG_CHI_BH (Mau 01)",
        "T_BHTT (Giam dinh XML)",
        "T_BHTT (Mau 01)",
    ]
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_xuat_excel.to_excel(writer, sheet_name="Tong_Hop", index=False)
        worksheet = writer.sheets["Tong_Hop"]

        colors = {
            "blue": PatternFill(start_color="B3D9FF", end_color="B3D9FF", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),
            "red": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
            "orange": PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid"),
            "purple": PatternFill(start_color="E6CCFF", end_color="E6CCFF", fill_type="solid"),
        }
        if "color" in df_tong_hop.columns:
            for row_idx, mau_to in enumerate(df_tong_hop["color"]):
                excel_row = row_idx + 2
                fill = colors.get(mau_to, None)
                if fill:
                    for col_idx in range(1, len(df_xuat_excel.columns) + 1):
                        worksheet.cell(row=excel_row, column=col_idx).fill = fill

        for col in worksheet.columns:
            ten_cot = col[0].value
            if ten_cot in ["STT"]:
                chieu_rong = 8
            elif ten_cot in cot_tien:
                chieu_rong = 16
            elif ten_cot in ["MA_THE (Giam dinh XML)", "MA_THE (Mau 01)"]:
                chieu_rong = 18
            elif ten_cot == "PHAN_LOAI":
                chieu_rong = 42
            else:
                chieu_rong = 20
            worksheet.column_dimensions[col[0].column_letter].width = chieu_rong
            if ten_cot in cot_tien:
                for cell in col[1:]:
                    cell.number_format = "#,##0.00"

    return file_xuat
