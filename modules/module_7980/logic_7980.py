import os
import re

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import chuyen_doi_ngay_7980, doc_file_thong_minh, lam_sach_chuoi


def safe_float(val):
    try:
        if pd.isna(val) or val == "":
            return 0.0
        clean_val = str(val).replace(",", "").strip()
        return float(clean_val)
    except Exception:
        return 0.0


def tim_cot(df, *ten_uu_tien):
    for ten in ten_uu_tien:
        if ten in df.columns:
            return ten
    cot_chuan = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}
    for ten in ten_uu_tien:
        ten_chuan = re.sub(r"\s+", " ", ten.strip().lower())
        if ten_chuan in cot_chuan:
            return cot_chuan[ten_chuan]
    raise KeyError(f"Khong tim thay cot nao trong danh sach: {ten_uu_tien}")


def parse_datetime_flexible(val):
    s = lam_sach_chuoi(val)
    if not s:
        return pd.NaT
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


def tao_key_dt(dt_val):
    if pd.isna(dt_val):
        return ""
    return pd.Timestamp(dt_val).strftime("%Y-%m-%d %H:%M")


def chuan_hoa_ma_the(val):
    raw = lam_sach_chuoi(val).upper()
    return re.sub(r"[^A-Z0-9]", "", raw)


def kiem_tra_cung_thang(thang_qt, ngay_tt):
    t_qt = str(thang_qt).strip()
    n_tt = str(ngay_tt).strip()
    if not t_qt or not n_tt or t_qt.lower() == "nan" or n_tt.lower() == "nan":
        return False
    t_qt = re.sub(r"\.0+$", "", t_qt)
    n_tt = re.sub(r"\.0+$", "", n_tt)

    thang_1, nam_1 = "", ""
    thang_2, nam_2 = "", ""

    if re.match(r"^\d{1,2}$", t_qt):
        thang_1 = t_qt.zfill(2)
    elif match := re.search(r"^(\d{4})(\d{2})$", t_qt):
        nam_1, thang_1 = match.group(1), match.group(2)
    elif match := re.search(r"(\d{1,2})[/_-](\d{4})", t_qt):
        thang_1, nam_1 = match.group(1).zfill(2), match.group(2)
    elif match := re.search(r"(\d{4})[/_-](\d{1,2})", t_qt):
        nam_1, thang_1 = match.group(1), match.group(2).zfill(2)

    if match := re.search(r"^(\d{4})(\d{2})\d{2}", n_tt):
        nam_2, thang_2 = match.group(1), match.group(2)
    elif match := re.search(r"\d{1,2}[/_-](\d{1,2})[/_-](\d{4})", n_tt):
        thang_2, nam_2 = match.group(1).zfill(2), match.group(2)
    elif match := re.search(r"(\d{4})[/_-](\d{1,2})[/_-](\d{1,2})", n_tt):
        nam_2, thang_2 = match.group(1), match.group(2).zfill(2)

    if not thang_1 or not thang_2:
        return False
    if nam_1:
        return (thang_1 == thang_2) and (nam_1 == nam_2)
    return thang_1 == thang_2


def tao_du_lieu_doi_chieu_7980(file_79_80, file_bh, tu_ngay, den_ngay, callback_status):
    try:
        callback_status("Dang doc du lieu tu Excel...", "blue")
        df_7980 = doc_file_thong_minh(file_79_80)
        df_bh = doc_file_thong_minh(file_bh)

        c7980_ma_bn = tim_cot(df_7980, "MA_BN")
        c7980_ho_ten = tim_cot(df_7980, "HO_TEN")
        c7980_ngay_vao = tim_cot(df_7980, "NGAY_VAO")
        c7980_ngay_ra = tim_cot(df_7980, "NGAY_RA")
        c7980_thang_qt = tim_cot(df_7980, "THANG_QT")
        c7980_tong_chi = tim_cot(df_7980, "T_TONGCHI")
        c7980_bhtt = tim_cot(df_7980, "T_BHTT")
        c7980_ma_the = tim_cot(df_7980, "MA_THE")

        cbh_ma_bn = tim_cot(df_bh, "Mã BN", "MA_BN", "Ma BN")
        cbh_ho_ten = tim_cot(df_bh, "Họ tên", "HO_TEN", "Ho ten")
        cbh_ngay_vao = tim_cot(df_bh, "Ngày vào", "NGAY_VAO", "Ngay vao")
        cbh_ngay_ra = tim_cot(df_bh, "Ngày ra", "NGAY_RA", "Ngay ra")
        cbh_ngay_tt = tim_cot(df_bh, "Ngày TT", "NGAY_TT", "Ngay TT")
        cbh_tong_chi = tim_cot(df_bh, "Tổng chi", "TONG_CHI", "Tong chi")
        cbh_bhtt = tim_cot(df_bh, "Bảo hiểm TT", "Bao hiem TT", "T_BHTT", "BAOHIEMTT", "TBHTT")
        cbh_ma_the = tim_cot(df_bh, "Mã thẻ", "MA_THE", "Ma the")

        df_7980 = df_7980.rename(columns={c7980_ma_bn: "MA_BN", c7980_ho_ten: "HO_TEN", c7980_ngay_vao: "NGAY_VAO", c7980_ngay_ra: "NGAY_RA", c7980_thang_qt: "THANG_QT", c7980_tong_chi: "T_TONGCHI", c7980_bhtt: "T_BHTT", c7980_ma_the: "MA_THE"})
        df_bh = df_bh.rename(columns={cbh_ma_bn: "MA_BN_BH", cbh_ho_ten: "HO_TEN_BH", cbh_ngay_vao: "NGAY_VAO_BH", cbh_ngay_ra: "NGAY_RA_BH", cbh_ngay_tt: "NGAY_TT_BH", cbh_tong_chi: "TONG_CHI_BH", cbh_bhtt: "T_BHTT_BH", cbh_ma_the: "MA_THE_BH"})

        callback_status("Dang chuan hoa du lieu...", "blue")
        df_7980["MA_BN"] = df_7980["MA_BN"].apply(lam_sach_chuoi)
        df_bh["MA_BN_BH"] = df_bh["MA_BN_BH"].apply(lam_sach_chuoi)
        df_7980["HO_TEN"] = df_7980["HO_TEN"].apply(lam_sach_chuoi)
        df_bh["HO_TEN_BH"] = df_bh["HO_TEN_BH"].apply(lam_sach_chuoi)
        df_7980["MA_THE"] = df_7980["MA_THE"].apply(lam_sach_chuoi)
        df_bh["MA_THE_BH"] = df_bh["MA_THE_BH"].apply(lam_sach_chuoi)
        df_7980["MA_THE_CHUAN"] = df_7980["MA_THE"].apply(chuan_hoa_ma_the)
        df_bh["MA_THE_BH_CHUAN"] = df_bh["MA_THE_BH"].apply(chuan_hoa_ma_the)

        df_7980["NGAY_VAO_CHUAN"] = df_7980["NGAY_VAO"].apply(chuyen_doi_ngay_7980)
        df_7980["NGAY_RA_CHUAN"] = df_7980["NGAY_RA"].apply(chuyen_doi_ngay_7980)
        df_7980["Ngay_Vao_DT_7980"] = df_7980["NGAY_VAO_CHUAN"].apply(parse_datetime_flexible)
        df_7980["Ngay_Ra_DT_7980"] = df_7980["NGAY_RA_CHUAN"].apply(parse_datetime_flexible)
        df_bh["Ngay_Vao_DT_BH"] = df_bh["NGAY_VAO_BH"].apply(parse_datetime_flexible)
        df_bh["Ngay_Ra_DT_BH"] = df_bh["NGAY_RA_BH"].apply(parse_datetime_flexible)
        df_bh["NGAY_VAO_CHUAN_BH"] = df_bh["Ngay_Vao_DT_BH"].apply(dinh_dang_dt_hien_thi)
        df_bh["NGAY_RA_CHUAN_BH"] = df_bh["Ngay_Ra_DT_BH"].apply(dinh_dang_dt_hien_thi)

        df_7980_valid = df_7980[df_7980["MA_BN"] != ""].copy()
        df_bh_valid = df_bh[df_bh["MA_BN_BH"] != ""].copy()

        callback_status("Dang loc theo khoang thoi gian...", "blue")
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        df_7980_valid = df_7980_valid.loc[(df_7980_valid["Ngay_Ra_DT_7980"] >= tu_ngay_dt) & (df_7980_valid["Ngay_Ra_DT_7980"] <= den_ngay_dt)].copy()
        df_bh_valid = df_bh_valid.loc[(df_bh_valid["Ngay_Ra_DT_BH"] >= tu_ngay_dt) & (df_bh_valid["Ngay_Ra_DT_BH"] <= den_ngay_dt)].copy()

        callback_status("Dang doi chieu du lieu...", "blue")
        df_7980_valid["KEY_NGAY_VAO"] = df_7980_valid["Ngay_Vao_DT_7980"].apply(tao_key_dt)
        df_7980_valid["KEY_NGAY_RA"] = df_7980_valid["Ngay_Ra_DT_7980"].apply(tao_key_dt)
        df_bh_valid["KEY_NGAY_VAO"] = df_bh_valid["Ngay_Vao_DT_BH"].apply(tao_key_dt)
        df_bh_valid["KEY_NGAY_RA"] = df_bh_valid["Ngay_Ra_DT_BH"].apply(tao_key_dt)
        df_7980_valid = df_7980_valid.sort_values(by=["MA_BN", "Ngay_Vao_DT_7980", "Ngay_Ra_DT_7980"])
        df_bh_valid = df_bh_valid.sort_values(by=["MA_BN_BH", "Ngay_Vao_DT_BH", "Ngay_Ra_DT_BH"])
        df_7980_valid["LAN_KHAM"] = df_7980_valid.groupby(["MA_BN", "KEY_NGAY_VAO"]).cumcount()
        df_bh_valid["LAN_KHAM"] = df_bh_valid.groupby(["MA_BN_BH", "KEY_NGAY_VAO"]).cumcount()

        merged = pd.merge(
            df_7980_valid,
            df_bh_valid,
            left_on=["MA_BN", "KEY_NGAY_VAO", "LAN_KHAM"],
            right_on=["MA_BN_BH", "KEY_NGAY_VAO", "LAN_KHAM"],
            how="outer",
            suffixes=("_7980", "_BH"),
        )

        rows = []
        for _, row in merged.iterrows():
            ton_tai_7980 = pd.notna(row["MA_BN"])
            ton_tai_bh = pd.notna(row["MA_BN_BH"])
            ma_bn_hien_thi = lam_sach_chuoi(row["MA_BN"]) if ton_tai_7980 else lam_sach_chuoi(row["MA_BN_BH"])
            thang_qt = str(row.get("THANG_QT", ""))
            ngay_tt = str(row.get("NGAY_TT_BH", ""))
            tien_7980 = safe_float(row.get("T_TONGCHI", 0))
            tien_bh = safe_float(row.get("TONG_CHI_BH", 0))
            bhtt_7980 = safe_float(row.get("T_BHTT", 0))
            bhtt_bh = safe_float(row.get("T_BHTT_BH", 0))
            ma_the_7980 = lam_sach_chuoi(row.get("MA_THE", ""))
            ma_the_bh = lam_sach_chuoi(row.get("MA_THE_BH", ""))
            ma_the_7980_chuan = lam_sach_chuoi(row.get("MA_THE_CHUAN", ""))
            ma_the_bh_chuan = lam_sach_chuoi(row.get("MA_THE_BH_CHUAN", ""))
            lech_tien = []
            if abs(tien_7980 - tien_bh) > 1.0:
                lech_tien.append("T_TONGCHI")
            if abs(bhtt_7980 - bhtt_bh) > 1.0:
                lech_tien.append("T_BHTT")

            phan_loai, mau_to = "Trùng khớp", "none"
            if ton_tai_7980 and not ton_tai_bh:
                phan_loai, mau_to = "Có trong 79_80 nhưng không có trong BH", "yellow"
            elif ton_tai_bh and not ton_tai_7980:
                phan_loai, mau_to = "Có trong BH nhưng không có trong 79_80", "blue"
            elif ma_the_7980_chuan != ma_the_bh_chuan:
                phan_loai, mau_to = "Lệch mã thẻ", "purple"
            elif lech_tien:
                phan_loai, mau_to = f"Sai lệch tiền: {', '.join(lech_tien)}", "red"
            elif (
                lam_sach_chuoi(row["KEY_NGAY_VAO"]) == ""
                or lam_sach_chuoi(row["NGAY_VAO_CHUAN"]) != lam_sach_chuoi(row["NGAY_VAO_CHUAN_BH"])
                or lam_sach_chuoi(row["NGAY_RA_CHUAN"]) != lam_sach_chuoi(row["NGAY_RA_CHUAN_BH"])
            ):
                phan_loai, mau_to = "Sai lệch Ngày vào/ra", "gray"
            elif not kiem_tra_cung_thang(thang_qt, ngay_tt):
                phan_loai, mau_to = "Ngày TT lệch Tháng QT", "orange"

            rows.append(
                {
                    "MÃ BỆNH NHÂN": ma_bn_hien_thi,
                    "TÊN BỆNH NHÂN": row["HO_TEN"] if ton_tai_7980 else row["HO_TEN_BH"],
                    "SỐ LƯỢNG": 1,
                    "NGÀY VÀO (79_80)": lam_sach_chuoi(row["NGAY_VAO_CHUAN"]),
                    "NGÀY VÀO (BH)": lam_sach_chuoi(row["NGAY_VAO_CHUAN_BH"]),
                    "NGÀY RA (79_80)": lam_sach_chuoi(row["NGAY_RA_CHUAN"]),
                    "NGÀY RA (BH)": lam_sach_chuoi(row["NGAY_RA_CHUAN_BH"]),
                    "THÁNG QUYẾT TOÁN (79_80)": thang_qt,
                    "NGÀY THANH TOÁN (BH)": ngay_tt,
                    "TỔNG CHI (79_80)": tien_7980,
                    "TỔNG CHI (BH)": tien_bh,
                    "T_BHTT (79_80)": bhtt_7980,
                    "Bảo hiểm TT (BH)": bhtt_bh,
                    "MÃ THẺ (79_80)": ma_the_7980,
                    "MÃ THẺ (BH)": ma_the_bh,
                    "PHÂN LOẠI": phan_loai,
                    "color": mau_to,
                }
            )

        df_tong_hop = pd.DataFrame(rows)
        thu_tu = [
            "Có trong BH nhưng không có trong 79_80",
            "Có trong 79_80 nhưng không có trong BH",
            "Lệch mã thẻ",
            "Sai lệch tiền",
            "Sai lệch Ngày vào/ra",
            "Ngày TT lệch Tháng QT",
            "Trùng khớp",
        ]
        df_tong_hop["_THU_TU"] = df_tong_hop["PHÂN LOẠI"].apply(
            lambda val: thu_tu.index("Sai lệch tiền") if str(val).startswith("Sai lệch tiền") else thu_tu.index(str(val)) if str(val) in thu_tu else 99
        )
        df_tong_hop = df_tong_hop.sort_values(["_THU_TU", "PHÂN LOẠI"]).drop(columns=["_THU_TU"]).reset_index(drop=True)
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
    cot_tien = ["TỔNG CHI (79_80)", "TỔNG CHI (BH)", "T_BHTT (79_80)", "Bảo hiểm TT (BH)"]
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_xuat_excel.to_excel(writer, sheet_name="Tong_Hop", index=False)
        worksheet = writer.sheets["Tong_Hop"]

        colors = {
            "blue": PatternFill(start_color="B3D9FF", end_color="B3D9FF", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),
            "red": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
            "orange": PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid"),
            "gray": PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"),
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
            if ten_cot in ["STT", "SỐ LƯỢNG"]:
                chieu_rong = 8
            elif ten_cot in cot_tien:
                chieu_rong = 16
            elif ten_cot in ["MÃ THẺ (79_80)", "MÃ THẺ (BH)"]:
                chieu_rong = 18
            elif ten_cot == "PHÂN LOẠI":
                chieu_rong = 40
            else:
                chieu_rong = 20
            worksheet.column_dimensions[col[0].column_letter].width = chieu_rong
            if ten_cot in cot_tien:
                for cell in col[1:]:
                    cell.number_format = "#,##0.00"

    return file_xuat


def xu_ly_doi_chieu(file_79_80, file_bh, file_xuat, tu_ngay, den_ngay, callback_status):
    thanh_cong, du_lieu_or_loi = tao_du_lieu_doi_chieu_7980(file_79_80, file_bh, tu_ngay, den_ngay, callback_status)
    if not thanh_cong:
        return False, du_lieu_or_loi
    try:
        callback_status("Dang xuat file Excel va to mau...", "blue")
        file_da_xuat = xuat_excel_tong_hop(du_lieu_or_loi, file_xuat)
        callback_status("Hoan thanh.", "green")
        try:
            os.startfile(file_da_xuat)
        except Exception:
            pass
        return True, file_da_xuat
    except Exception as e:
        callback_status("Loi xay ra.", "red")
        return False, str(e)
