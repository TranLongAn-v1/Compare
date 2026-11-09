import os
import re
import unicodedata
import warnings

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import chuyen_doi_ngay_7980, doc_file_thong_minh, lam_sach_chuoi

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
        return float(str(val).replace(",", "").strip())
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


def doc_file_mau01_thong_minh(file_path):
    raw = pd.read_excel(file_path, engine="openpyxl", header=None)
    nhom_bat_buoc = [
        {"mathe", "mathebhyt"},
        {"hoten", "hovaten"},
        {"ngayvao"},
        {"ngayra"},
        {"ttongchibh", "tongchibh"},
        {"ttongchibv", "tongchibv"},
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
        return doc_file_thong_minh(file_path)

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
    s = re.sub(r"\s+(SA|CH)\s*$", "", s, flags=re.IGNORECASE)
    if re.fullmatch(r"\d{12}", s):
        s = chuyen_doi_ngay_7980(s)
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


def tao_du_lieu_doi_chieu_xml01_bh_mau_chuan(file_xml01_bh, file_mau_chuan, tu_ngay, den_ngay, callback_status):
    try:
        callback_status("Dang doc du lieu tu Excel...", "blue")
        df_xml = doc_file_mau01_thong_minh(file_xml01_bh)
        df_chuan = doc_file_mau01_thong_minh(file_mau_chuan)

        cxml_ma_the = tim_cot(df_xml, "MA_THE_BHYT", "MA_THE", "Ma the", "MATHE")
        cxml_ho_ten = tim_cot(df_xml, "HO_TEN", "Ho ten", "Ho va ten", "HOTEN")
        cxml_ngay_vao = tim_cot(df_xml, "NGAY_VAO", "Ngay vao", "NGAYVAO")
        cxml_ngay_ra = tim_cot(df_xml, "NGAY_RA", "Ngay ra", "NGAYRA")
        cxml_tong_bh = tim_cot(df_xml, "T_TONGCHI_BH", "TONG_CHI_BH", "Tong chi BH", "TONGCHIBH")
        cxml_tong_bv = tim_cot(df_xml, "T_TONGCHI_BV", "TONG_CHI_BV", "Tong chi BV", "TONGCHIBV")
        cxml_bhtt = tim_cot(df_xml, "T_BHTT", "Bao hiem TT", "BAOHIEMTT", "TBHTT")

        cchuan_ma_the = tim_cot(df_chuan, "MA_THE", "MA_THE_BHYT", "Ma the", "MATHE")
        cchuan_ho_ten = tim_cot(df_chuan, "HO_TEN", "Ho ten", "Ho va ten", "HOTEN")
        cchuan_ngay_vao = tim_cot(df_chuan, "NGAY_VAO", "Ngay vao", "NGAYVAO")
        cchuan_ngay_ra = tim_cot(df_chuan, "NGAY_RA", "Ngay ra", "NGAYRA")
        cchuan_tong_bh = tim_cot(df_chuan, "T_TONGCHI_BH", "TONG_CHI_BH", "Tong chi BH", "TONGCHIBH")
        cchuan_tong_bv = tim_cot(df_chuan, "T_TONGCHI_BV", "TONG_CHI_BV", "Tong chi BV", "TONGCHIBV")
        cchuan_bhtt = tim_cot(df_chuan, "T_BHTT", "Bao hiem TT", "BAOHIEMTT", "TBHTT")

        df_xml = df_xml.rename(
            columns={
                cxml_ma_the: "MA_THE_XML",
                cxml_ho_ten: "HO_TEN_XML",
                cxml_ngay_vao: "NGAY_VAO_XML",
                cxml_ngay_ra: "NGAY_RA_XML",
                cxml_tong_bh: "TONG_BH_XML",
                cxml_tong_bv: "TONG_BV_XML",
                cxml_bhtt: "T_BHTT_XML",
            }
        )
        df_chuan = df_chuan.rename(
            columns={
                cchuan_ma_the: "MA_THE_CHUAN_SRC",
                cchuan_ho_ten: "HO_TEN_CHUAN_SRC",
                cchuan_ngay_vao: "NGAY_VAO_CHUAN_SRC",
                cchuan_ngay_ra: "NGAY_RA_CHUAN_SRC",
                cchuan_tong_bh: "TONG_BH_CHUAN",
                cchuan_tong_bv: "TONG_BV_CHUAN",
                cchuan_bhtt: "T_BHTT_CHUAN",
            }
        )

        callback_status("Dang chuan hoa du lieu...", "blue")
        for col in ["MA_THE_XML", "HO_TEN_XML"]:
            df_xml[col] = df_xml[col].apply(lam_sach_chuoi)
        for col in ["MA_THE_CHUAN_SRC", "HO_TEN_CHUAN_SRC"]:
            df_chuan[col] = df_chuan[col].apply(lam_sach_chuoi)

        df_xml["MA_THE_CHUAN"] = df_xml["MA_THE_XML"].apply(chuan_hoa_ma_the)
        df_chuan["MA_THE_CHUAN"] = df_chuan["MA_THE_CHUAN_SRC"].apply(chuan_hoa_ma_the)
        df_xml["HO_TEN_KEY"] = df_xml["HO_TEN_XML"].apply(chuan_hoa_ho_ten)
        df_chuan["HO_TEN_KEY"] = df_chuan["HO_TEN_CHUAN_SRC"].apply(chuan_hoa_ho_ten)

        df_xml["Ngay_Vao_DT"] = df_xml["NGAY_VAO_XML"].apply(parse_datetime_flexible)
        df_xml["Ngay_Ra_DT"] = df_xml["NGAY_RA_XML"].apply(parse_datetime_flexible)
        df_chuan["Ngay_Vao_DT"] = df_chuan["NGAY_VAO_CHUAN_SRC"].apply(parse_datetime_flexible)
        df_chuan["Ngay_Ra_DT"] = df_chuan["NGAY_RA_CHUAN_SRC"].apply(parse_datetime_flexible)

        df_xml_valid = df_xml[(df_xml["HO_TEN_KEY"] != "")].copy()
        df_chuan_valid = df_chuan[(df_chuan["HO_TEN_KEY"] != "")].copy()

        callback_status("Dang loc theo khoang thoi gian...", "blue")
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        df_xml_valid = df_xml_valid.loc[
            (df_xml_valid["Ngay_Ra_DT"] >= tu_ngay_dt) & (df_xml_valid["Ngay_Ra_DT"] <= den_ngay_dt)
        ].copy()
        df_chuan_valid = df_chuan_valid.loc[
            (df_chuan_valid["Ngay_Ra_DT"] >= tu_ngay_dt) & (df_chuan_valid["Ngay_Ra_DT"] <= den_ngay_dt)
        ].copy()

        callback_status("Dang doi chieu du lieu...", "blue")
        rows = []
        ds_xml = df_xml_valid.reset_index(drop=True).to_dict("records")
        ds_chuan = df_chuan_valid.reset_index(drop=True).to_dict("records")
        da_ghep_xml = set()

        for row_chuan in ds_chuan:
            idx_tot_nhat = None
            diem_tot_nhat = None
            for idx, row_xml in enumerate(ds_xml):
                if idx in da_ghep_xml:
                    continue
                if row_xml.get("HO_TEN_KEY", "") != row_chuan.get("HO_TEN_KEY", ""):
                    continue

                cung_ma_the = row_xml.get("MA_THE_CHUAN", "") == row_chuan.get("MA_THE_CHUAN", "")
                cung_ngay = (
                    dinh_dang_dt_hien_thi(row_xml.get("Ngay_Vao_DT", pd.NaT)) == dinh_dang_dt_hien_thi(row_chuan.get("Ngay_Vao_DT", pd.NaT))
                    and dinh_dang_dt_hien_thi(row_xml.get("Ngay_Ra_DT", pd.NaT)) == dinh_dang_dt_hien_thi(row_chuan.get("Ngay_Ra_DT", pd.NaT))
                )
                muc_uu_tien = 1 if cung_ma_the else 2 if cung_ngay else 3
                diem = (muc_uu_tien, _so_sanh_datetime(row_xml.get("Ngay_Vao_DT", pd.NaT), row_chuan.get("Ngay_Vao_DT", pd.NaT)))
                if diem_tot_nhat is None or diem < diem_tot_nhat:
                    diem_tot_nhat = diem
                    idx_tot_nhat = idx

            if idx_tot_nhat is not None:
                da_ghep_xml.add(idx_tot_nhat)
                row_xml = ds_xml[idx_tot_nhat]
                ton_tai_xml = True
            else:
                row_xml = {}
                ton_tai_xml = False

            rows.append(_tao_dong_ket_qua(row_xml, row_chuan, ton_tai_xml, True))

        for idx, row_xml in enumerate(ds_xml):
            if idx in da_ghep_xml:
                continue
            rows.append(_tao_dong_ket_qua(row_xml, {}, True, False))

        df_tong_hop = pd.DataFrame(rows)
        if not df_tong_hop.empty:
            thu_tu = {
                "Co tren Mau chuan nhung khong co tren Xml 01-BH": 1,
                "Co tren Xml 01-BH nhung khong co tren Mau chuan": 2,
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


def _tao_dong_ket_qua(row_xml, row_chuan, ton_tai_xml, ton_tai_chuan):
    ngay_vao_xml = dinh_dang_dt_hien_thi(row_xml.get("Ngay_Vao_DT", pd.NaT))
    ngay_ra_xml = dinh_dang_dt_hien_thi(row_xml.get("Ngay_Ra_DT", pd.NaT))
    ngay_vao_chuan = dinh_dang_dt_hien_thi(row_chuan.get("Ngay_Vao_DT", pd.NaT))
    ngay_ra_chuan = dinh_dang_dt_hien_thi(row_chuan.get("Ngay_Ra_DT", pd.NaT))

    tong_bh_xml = safe_float(row_xml.get("TONG_BH_XML", 0))
    tong_bh_chuan = safe_float(row_chuan.get("TONG_BH_CHUAN", 0))
    tong_bv_xml = safe_float(row_xml.get("TONG_BV_XML", 0))
    tong_bv_chuan = safe_float(row_chuan.get("TONG_BV_CHUAN", 0))
    bhtt_xml = safe_float(row_xml.get("T_BHTT_XML", 0))
    bhtt_chuan = safe_float(row_chuan.get("T_BHTT_CHUAN", 0))
    ma_the_xml = lam_sach_chuoi(row_xml.get("MA_THE_XML", ""))
    ma_the_chuan = lam_sach_chuoi(row_chuan.get("MA_THE_CHUAN_SRC", ""))

    lech_tien = []
    if abs(tong_bh_xml - tong_bh_chuan) > 1.0:
        lech_tien.append("T_TONGCHI_BH")
    if abs(tong_bv_xml - tong_bv_chuan) > 1.0:
        lech_tien.append("T_TONGCHI_BV")
    if abs(bhtt_xml - bhtt_chuan) > 1.0:
        lech_tien.append("T_BHTT")

    phan_loai, mau_to = "Trung khop", "none"
    if ton_tai_chuan and not ton_tai_xml:
        phan_loai, mau_to = "Co tren Mau chuan nhung khong co tren Xml 01-BH", "blue"
    elif ton_tai_xml and not ton_tai_chuan:
        phan_loai, mau_to = "Co tren Xml 01-BH nhung khong co tren Mau chuan", "yellow"
    elif lam_sach_chuoi(row_xml.get("MA_THE_CHUAN", "")) != lam_sach_chuoi(row_chuan.get("MA_THE_CHUAN", "")):
        phan_loai, mau_to = "Lech ma the", "purple"
    elif (ngay_vao_xml != ngay_vao_chuan) or (ngay_ra_xml != ngay_ra_chuan):
        phan_loai, mau_to = "Lech ngay vao/ngay ra", "orange"
    elif lech_tien:
        phan_loai, mau_to = f"Lech tien: {', '.join(lech_tien)}", "red"

    return {
        "HO_TEN": row_chuan.get("HO_TEN_CHUAN_SRC", "") if ton_tai_chuan else row_xml.get("HO_TEN_XML", ""),
        "MA_THE (Xml 01-BH)": ma_the_xml,
        "MA_THE (Mau chuan)": ma_the_chuan,
        "NGAY_VAO (Xml 01-BH)": ngay_vao_xml,
        "NGAY_VAO (Mau chuan)": ngay_vao_chuan,
        "NGAY_RA (Xml 01-BH)": ngay_ra_xml,
        "NGAY_RA (Mau chuan)": ngay_ra_chuan,
        "T_TONGCHI_BH (Xml 01-BH)": tong_bh_xml,
        "T_TONGCHI_BH (Mau chuan)": tong_bh_chuan,
        "T_TONGCHI_BV (Xml 01-BH)": tong_bv_xml,
        "T_TONGCHI_BV (Mau chuan)": tong_bv_chuan,
        "T_BHTT (Xml 01-BH)": bhtt_xml,
        "T_BHTT (Mau chuan)": bhtt_chuan,
        "PHAN_LOAI": phan_loai,
        "color": mau_to,
    }


def xuat_excel_tong_hop(df_tong_hop, file_xuat):
    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    df_xuat_excel = df_tong_hop.drop(columns=["color"], errors="ignore")
    cot_tien = [
        "T_TONGCHI_BH (Xml 01-BH)",
        "T_TONGCHI_BH (Mau chuan)",
        "T_TONGCHI_BV (Xml 01-BH)",
        "T_TONGCHI_BV (Mau chuan)",
        "T_BHTT (Xml 01-BH)",
        "T_BHTT (Mau chuan)",
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
                fill = colors.get(mau_to, None)
                if fill:
                    for col_idx in range(1, len(df_xuat_excel.columns) + 1):
                        worksheet.cell(row=row_idx + 2, column=col_idx).fill = fill

        for col in worksheet.columns:
            ten_cot = col[0].value
            if ten_cot in ["STT"]:
                chieu_rong = 8
            elif ten_cot in cot_tien:
                chieu_rong = 18
            elif ten_cot in ["MA_THE (Xml 01-BH)", "MA_THE (Mau chuan)"]:
                chieu_rong = 20
            elif ten_cot == "PHAN_LOAI":
                chieu_rong = 48
            else:
                chieu_rong = 22
            worksheet.column_dimensions[col[0].column_letter].width = chieu_rong
            if ten_cot in cot_tien:
                for cell in col[1:]:
                    cell.number_format = "#,##0.00"

    return file_xuat
