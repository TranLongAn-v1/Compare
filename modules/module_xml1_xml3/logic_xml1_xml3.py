import os
import re

import pandas as pd

from core.utils import doc_file_thong_minh

MA_NHOM_LOAI_BO = {3, 8, 10, 13, 14, 15, 18}
COT_XML1_BAT_BUOC = ["MA_LK", "MA_HSBA", "MA_THE_BHYT", "MA_BN", "HO_TEN", "NGAY_VAO", "NGAY_RA"]
COT_XML3_BAT_BUOC = ["MA_LK", "MA_BA", "MA_THE", "MA_NHOM", "MA_DICH_VU", "TEN_DICH_VU", "MA_MAY"]
COT_KET_QUA = [
    "MA_HSBA",
    "MA_BN",
    "HO_TEN",
    "NGAY_VAO",
    "NGAY_RA",
    "MA_DICH_VU",
    "TEN_DICH_VU",
    "MA_MAY",
    "Khớp MA_LK",
    "Khớp MA_HSBA/MA_BA",
    "Khớp MA_THE_BHYT/MA_THE",
]


def _is_xml1_filename(path_file):
    ten_file = os.path.splitext(os.path.basename(path_file))[0]
    return re.match(r"^xml1", ten_file, flags=re.IGNORECASE) is not None


def _is_xml3_filename(path_file):
    ten_file = os.path.splitext(os.path.basename(path_file))[0]
    return re.match(r"^xml3", ten_file, flags=re.IGNORECASE) is not None


def _as_key_series(series):
    return series.fillna("").astype(str).str.strip()


def _dinh_dang_yyyymmddhhmm(val):
    s = str(val).strip()
    if len(s) == 12 and s.isdigit():
        return f"{s[6:8]}/{s[4:6]}/{s[0:4]} {s[8:10]}:{s[10:12]}"
    return s


def _chuan_hoa_ma_nhom(val):
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def doc_kiem_tra_xml1(path_file):
    if not _is_xml1_filename(path_file):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    thieu = [c for c in COT_XML1_BAT_BUOC if c not in df.columns]
    if thieu:
        raise ValueError("File không hợp lệ")
    return df


def doc_kiem_tra_xml3(path_file):
    if not _is_xml3_filename(path_file):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    thieu = [c for c in COT_XML3_BAT_BUOC if c not in df.columns]
    if thieu:
        raise ValueError("File không hợp lệ")
    return df


def ket_xuat_xml1_xml3(path_xml1, path_xml3, callback_status=None):
    if callback_status:
        callback_status("Đang đọc file XML1...", "blue")
    df1 = doc_kiem_tra_xml1(path_xml1).copy()

    if callback_status:
        callback_status("Đang đọc file XML3...", "blue")
    df3 = doc_kiem_tra_xml3(path_xml3).copy()

    for cot in ["MA_LK", "MA_HSBA", "MA_THE_BHYT", "MA_BN", "HO_TEN", "NGAY_VAO", "NGAY_RA"]:
        df1[cot] = _as_key_series(df1[cot])

    for cot in ["MA_LK", "MA_BA", "MA_THE", "MA_DICH_VU", "TEN_DICH_VU", "MA_MAY"]:
        df3[cot] = _as_key_series(df3[cot])

    # Đối chiếu định danh trên XML3 raw
    tap_ma_lk_xml3 = set(df3["MA_LK"].tolist())
    df_id = df3[["MA_LK", "MA_BA", "MA_THE"]].drop_duplicates()
    df_danh_dau = df1.merge(
        df_id,
        how="left",
        left_on=["MA_LK", "MA_HSBA"],
        right_on=["MA_LK", "MA_BA"],
    )

    # Lọc MA_NHOM chỉ để lấy cột dịch vụ hiển thị
    if callback_status:
        callback_status("Đang lọc XML3 theo MA_NHOM để lấy dịch vụ...", "blue")
    ma_nhom_series = df3["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    df3_dich_vu = df3.loc[~ma_nhom_series.isin(MA_NHOM_LOAI_BO)].copy()

    df_kq = df1.merge(
        df3_dich_vu[["MA_LK", "MA_BA", "MA_DICH_VU", "TEN_DICH_VU", "MA_MAY"]],
        how="left",
        left_on=["MA_LK", "MA_HSBA"],
        right_on=["MA_LK", "MA_BA"],
    )

    key_matched = set(zip(df_danh_dau["MA_LK"], df_danh_dau["MA_BA"]))
    df_kq["Khớp MA_LK"] = df_kq["MA_LK"].apply(lambda v: "Có" if v in tap_ma_lk_xml3 else "Không")
    df_kq["Khớp MA_HSBA/MA_BA"] = df_kq.apply(
        lambda r: "Có" if (r["MA_LK"], r["MA_HSBA"]) in key_matched else "Không",
        axis=1,
    )

    the_by_key = {}
    for _, r in df_id.iterrows():
        key = (r["MA_LK"], r["MA_BA"])
        the = str(r["MA_THE"]).strip()
        if not the:
            continue
        if key not in the_by_key:
            the_by_key[key] = set()
        the_by_key[key].add(the)

    def _khop_the(row):
        ma_the_bhyt = str(row["MA_THE_BHYT"]).strip()
        if not ma_the_bhyt:
            return "Không"
        tap_the = the_by_key.get((row["MA_LK"], row["MA_HSBA"]), set())
        return "Có" if ma_the_bhyt in tap_the else "Không"

    df_kq["Khớp MA_THE_BHYT/MA_THE"] = df_kq.apply(_khop_the, axis=1)

    for cot_ngay in ["NGAY_VAO", "NGAY_RA"]:
        df_kq[cot_ngay] = df_kq[cot_ngay].apply(_dinh_dang_yyyymmddhhmm)

    df_kq = df_kq[COT_KET_QUA].copy().fillna("")

    if callback_status:
        callback_status(f"Kết xuất xong: {len(df_kq)} dòng.", "green")

    return df_kq.reset_index(drop=True)


def xuat_excel_xml1_xml3(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    base, ext = os.path.splitext(file_xuat)
    idx = 1
    file_thuc_te = file_xuat
    while os.path.exists(file_thuc_te):
        file_thuc_te = f"{base} ({idx}){ext}"
        idx += 1

    with pd.ExcelWriter(file_thuc_te, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="XuLy_XML1_XML3", index=False)
        ws = writer.sheets["XuLy_XML1_XML3"]
        for col in ws.columns:
            ten = str(col[0].value)
            if ten in {"HO_TEN", "TEN_DICH_VU"}:
                width = 30
            elif ten in {"MA_HSBA", "MA_BN", "MA_DICH_VU", "MA_MAY"}:
                width = 18
            elif ten.startswith("Khớp"):
                width = 20
            else:
                width = 16
            ws.column_dimensions[col[0].column_letter].width = width

    return file_thuc_te
