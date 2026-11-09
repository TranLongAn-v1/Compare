# modules/module_danh_sach_thay_the/logic_danh_sach_thay_the.py
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
    raise KeyError(f"Không tìm thấy cột nào trong danh sách: {ten_uu_tien}")


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


def tao_du_lieu_doi_chieu_danh_sach_thay_the(file_dstt, file_bh, tu_ngay, den_ngay, callback_status):
    try:
        callback_status("Đang đọc dữ liệu từ Excel...", "blue")
        df_dstt = doc_file_thong_minh(file_dstt)
        df_bh = doc_file_thong_minh(file_bh)

        cdstt_ma_bn = tim_cot(df_dstt, "MA_BN", "MÃ BỆNH NHÂN", "Ma benh nhan")
        cdstt_ho_ten = tim_cot(df_dstt, "HO_TEN", "TÊN BỆNH NHÂN", "Ten benh nhan")
        cdstt_ngay_vao = tim_cot(df_dstt, "NGAY_VAO", "NGÀY VÀO (79_80)", "Ngay vao (79_80)", "NGÀY VÀO")
        cdstt_ngay_ra = tim_cot(df_dstt, "NGAY_RA", "NGÀY RA (79_80)", "Ngay ra (79_80)", "NGÀY RA")
        cdstt_thang_qt = tim_cot(df_dstt, "THANG_QT", "THÁNG QUYẾT TOÁN (79_80)", "Thang quyet toan (79_80)", "THÁNG QUYẾT TOÁN")
        cdstt_tong_chi = tim_cot(df_dstt, "T_TONGCHI", "TỔNG CHI (79_80)", "Tong chi (79_80)", "TỔNG CHI")
        cdstt_ma_the = tim_cot(df_dstt, "MA_THE", "MÃ THẺ (79_80)", "Ma the (79_80)", "MÃ THẺ")

        cbh_ma_bn = tim_cot(df_bh, "Mã BN", "MA_BN", "Ma BN")
        cbh_ho_ten = tim_cot(df_bh, "Họ tên", "HO_TEN", "Ho ten")
        cbh_ngay_vao = tim_cot(df_bh, "Ngày vào", "NGAY_VAO", "Ngay vao")
        cbh_ngay_ra = tim_cot(df_bh, "Ngày ra", "NGAY_RA", "Ngay ra")
        cbh_ngay_tt = tim_cot(df_bh, "Ngày TT", "NGAY_TT", "Ngay TT")
        cbh_tong_chi = tim_cot(df_bh, "Tổng chi", "TONG_CHI", "Tong chi")
        cbh_ma_the = tim_cot(df_bh, "Mã thẻ", "MA_THE", "Ma the")

        df_dstt = df_dstt.rename(columns={cdstt_ma_bn: "MA_BN", cdstt_ho_ten: "HO_TEN", cdstt_ngay_vao: "NGAY_VAO", cdstt_ngay_ra: "NGAY_RA", cdstt_thang_qt: "THANG_QT", cdstt_tong_chi: "T_TONGCHI", cdstt_ma_the: "MA_THE"})
        df_bh = df_bh.rename(columns={cbh_ma_bn: "MA_BN_BH", cbh_ho_ten: "HO_TEN_BH", cbh_ngay_vao: "NGAY_VAO_BH", cbh_ngay_ra: "NGAY_RA_BH", cbh_ngay_tt: "NGAY_TT_BH", cbh_tong_chi: "TONG_CHI_BH", cbh_ma_the: "MA_THE_BH"})

        callback_status("Đang chuẩn hóa dữ liệu...", "blue")
        df_dstt["MA_BN"] = df_dstt["MA_BN"].apply(lam_sach_chuoi)
        df_bh["MA_BN_BH"] = df_bh["MA_BN_BH"].apply(lam_sach_chuoi)
        df_dstt["HO_TEN"] = df_dstt["HO_TEN"].apply(lam_sach_chuoi)
        df_bh["HO_TEN_BH"] = df_bh["HO_TEN_BH"].apply(lam_sach_chuoi)
        df_dstt["MA_THE"] = df_dstt["MA_THE"].apply(lam_sach_chuoi)
        df_bh["MA_THE_BH"] = df_bh["MA_THE_BH"].apply(lam_sach_chuoi)
        df_dstt["MA_THE_CHUAN"] = df_dstt["MA_THE"].apply(chuan_hoa_ma_the)
        df_bh["MA_THE_BH_CHUAN"] = df_bh["MA_THE_BH"].apply(chuan_hoa_ma_the)

        df_dstt["NGAY_VAO_CHUAN"] = df_dstt["NGAY_VAO"].apply(chuyen_doi_ngay_7980)
        df_dstt["NGAY_RA_CHUAN"] = df_dstt["NGAY_RA"].apply(chuyen_doi_ngay_7980)
        df_dstt["Ngay_Vao_DT_7980"] = df_dstt["NGAY_VAO_CHUAN"].apply(parse_datetime_flexible)
        df_dstt["Ngay_Ra_DT_7980"] = df_dstt["NGAY_RA_CHUAN"].apply(parse_datetime_flexible)
        df_bh["Ngay_Vao_DT_BH"] = df_bh["NGAY_VAO_BH"].apply(parse_datetime_flexible)
        df_bh["Ngay_Ra_DT_BH"] = df_bh["NGAY_RA_BH"].apply(parse_datetime_flexible)
        df_bh["NGAY_VAO_CHUAN_BH"] = df_bh["Ngay_Vao_DT_BH"].apply(dinh_dang_dt_hien_thi)
        df_bh["NGAY_RA_CHUAN_BH"] = df_bh["Ngay_Ra_DT_BH"].apply(dinh_dang_dt_hien_thi)

        df_dstt_valid = df_dstt[df_dstt["MA_BN"] != ""].copy()
        df_bh_valid = df_bh[df_bh["MA_BN_BH"] != ""].copy()

        callback_status("Đang lọc theo khoảng thời gian...", "blue")
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        df_dstt_valid = df_dstt_valid.loc[(df_dstt_valid["Ngay_Ra_DT_7980"] >= tu_ngay_dt) & (df_dstt_valid["Ngay_Ra_DT_7980"] <= den_ngay_dt)].copy()
        df_bh_valid = df_bh_valid.loc[(df_bh_valid["Ngay_Ra_DT_BH"] >= tu_ngay_dt) & (df_bh_valid["Ngay_Ra_DT_BH"] <= den_ngay_dt)].copy()

        callback_status("Đang đối chiếu dữ liệu...", "blue")
        df_dstt_valid["KEY_NGAY_VAO"] = df_dstt_valid["Ngay_Vao_DT_7980"].apply(tao_key_dt)
        df_dstt_valid["KEY_NGAY_RA"] = df_dstt_valid["Ngay_Ra_DT_7980"].apply(tao_key_dt)
        df_bh_valid["KEY_NGAY_VAO"] = df_bh_valid["Ngay_Vao_DT_BH"].apply(tao_key_dt)
        df_bh_valid["KEY_NGAY_RA"] = df_bh_valid["Ngay_Ra_DT_BH"].apply(tao_key_dt)
        df_dstt_valid = df_dstt_valid.sort_values(by=["MA_BN", "Ngay_Vao_DT_7980", "Ngay_Ra_DT_7980"])
        df_bh_valid = df_bh_valid.sort_values(by=["MA_BN_BH", "Ngay_Vao_DT_BH", "Ngay_Ra_DT_BH"])
        df_dstt_valid["LAN_KHAM"] = df_dstt_valid.groupby(["MA_BN", "KEY_NGAY_VAO"]).cumcount()
        df_bh_valid["LAN_KHAM"] = df_bh_valid.groupby(["MA_BN_BH", "KEY_NGAY_VAO"]).cumcount()

        merged = pd.merge(
            df_dstt_valid,
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
            ma_the_7980 = lam_sach_chuoi(row.get("MA_THE", ""))
            ma_the_bh = lam_sach_chuoi(row.get("MA_THE_BH", ""))

            phan_loai, mau_to = "Trùng khớp thay thế", "none"
            if ton_tai_7980 and not ton_tai_bh:
                phan_loai, mau_to = "Không tồn tại", "yellow"
            elif ton_tai_bh and not ton_tai_7980:
                phan_loai, mau_to = "Không tồn tại", "blue"
            elif abs(tien_7980 - tien_bh) > 1.0:
                phan_loai, mau_to = "Sai lệch tổng chi", "red"
            elif (
                not kiem_tra_cung_thang(thang_qt, ngay_tt)
                or lam_sach_chuoi(row["KEY_NGAY_VAO"]) == ""
                or lam_sach_chuoi(row["NGAY_VAO_CHUAN"]) != lam_sach_chuoi(row["NGAY_VAO_CHUAN_BH"])
                or lam_sach_chuoi(row["NGAY_RA_CHUAN"]) != lam_sach_chuoi(row["NGAY_RA_CHUAN_BH"])
            ):
                phan_loai, mau_to = "Sai lệch ngày vào ngày ra", "orange"

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
                    "MÃ THẺ (79_80)": ma_the_7980,
                    "MÃ THẺ (BH)": ma_the_bh,
                    "PHÂN LOẠI": phan_loai,
                    "color": mau_to,
                }
            )

        df_tong_hop = pd.DataFrame(rows)
        thu_tu = [
            "Không tồn tại",
            "Sai lệch tổng chi",
            "Sai lệch ngày vào ngày ra",
            "Trùng khớp thay thế",
        ]
        df_tong_hop["PHÂN LOẠI"] = pd.Categorical(df_tong_hop["PHÂN LOẠI"], categories=thu_tu, ordered=True)
        df_tong_hop = df_tong_hop.sort_values("PHÂN LOẠI").reset_index(drop=True)
        df_tong_hop.insert(0, "STT", range(1, len(df_tong_hop) + 1))
        callback_status("Hoàn thành đối chiếu.", "green")
        return True, df_tong_hop
    except Exception as e:
        callback_status("Lỗi xảy ra.", "red")
        return False, str(e)


def xuat_excel_tong_hop_dstt(df_tong_hop, file_xuat):
    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    df_xuat_excel = df_tong_hop.drop(columns=["color"], errors="ignore")
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_xuat_excel.to_excel(writer, sheet_name="Tong_Hop", index=False)
        worksheet = writer.sheets["Tong_Hop"]

        colors = {
            "blue": PatternFill(start_color="B3D9FF", end_color="B3D9FF", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),
            "red": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
            "orange": PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid"),
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
            elif ten_cot in ["TỔNG CHI (79_80)", "TỔNG CHI (BH)"]:
                chieu_rong = 16
            elif ten_cot in ["MÃ THẺ (79_80)", "MÃ THẺ (BH)"]:
                chieu_rong = 18
            elif ten_cot == "PHÂN LOẠI":
                chieu_rong = 40
            else:
                chieu_rong = 20
            worksheet.column_dimensions[col[0].column_letter].width = chieu_rong
            if ten_cot in ["TỔNG CHI (79_80)", "TỔNG CHI (BH)"]:
                for cell in col[1:]:
                    cell.number_format = "#,##0.00"

    return file_xuat
