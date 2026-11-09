import os

import pandas as pd

from core.utils import doc_file_thong_minh


COT_XML3_CHUAN = ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY", "MA_BA", "HO_TEN"]
COT_ANH_XA_BAT_BUOC = [
    "STT",
    "TEN_TB",
    "KY_HIEU",
    "CONGTY_SX",
    "NUOC_SX",
    "NAM_SX",
    "NAM_SD",
    "MA_MAY",
    "SO_LUU_HANH",
    "HD_TU",
    "HD_DEN",
    "TU_NGAY",
    "DEN_NGAY",
    "ID",
]


def _ten_file_khong_ext(path_file):
    return os.path.splitext(os.path.basename(path_file))[0]


def _series_chuan_hoa(series):
    return series.fillna("").astype(str).str.strip()


def _series_ma_may_chuan_hoa(series):
    s = series.fillna("").astype(str).str.strip()
    return s.str.replace(r"\.0+$", "", regex=True)


def doc_xml3_tong_hop(path_file):
    ten = _ten_file_khong_ext(path_file).lower()
    if not ten.startswith("tong_hop_xml3"):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    cot_thuc_te = [str(c).strip() for c in df.columns.tolist()]
    if set(cot_thuc_te) != set(COT_XML3_CHUAN):
        raise ValueError("File không hợp lệ")

    return df


def doc_dm_may(path_file):
    ten = _ten_file_khong_ext(path_file).lower()
    if not ten.startswith("dm_may"):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    cot = {str(c).strip() for c in df.columns.tolist()}
    if "MAMAY3176" not in cot or "MAMAY" not in cot:
        raise ValueError("File không hợp lệ")

    return df


def doc_file_anh_xa(path_file):
    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    cot_thuc_te = {str(c).strip() for c in df.columns.tolist()}
    if not set(COT_ANH_XA_BAT_BUOC).issubset(cot_thuc_te):
        raise ValueError("File không hợp lệ")

    return df


def doi_chieu_ma_may(df_xml3, df_dm):
    df = df_xml3.copy()

    dm_3176 = set(_series_chuan_hoa(df_dm["MAMAY3176"]))
    dm_mamay = set(_series_chuan_hoa(df_dm["MAMAY"]))

    ma_may = _series_chuan_hoa(df["MA_MAY"])
    khop_3176 = ma_may.isin(dm_3176)
    khop_mamay = ma_may.isin(dm_mamay)

    df["Khớp MAMAY3176"] = khop_3176.map(lambda x: "Có" if x else "Không")
    df["Khớp MAMAY"] = khop_mamay.map(lambda x: "Có" if x else "Không")
    df["Không tồn tại trong DM"] = (~(khop_3176 | khop_mamay)).map(lambda x: "Có" if x else "Không")

    cot_hien_thi = [
        "MA_DICH_VU",
        "TEN_DICH_VU",
        "MA_MAY",
        "MA_BA",
        "HO_TEN",
        "Khớp MAMAY3176",
        "Khớp MAMAY",
        "Không tồn tại trong DM",
    ]
    return df[cot_hien_thi].reset_index(drop=True)


def doi_chieu_anh_xa(df_dm, df_anh_xa):
    df = df_anh_xa.copy()

    dm_3176 = set(_series_ma_may_chuan_hoa(df_dm["MAMAY3176"]))
    ma_may_anh_xa = _series_ma_may_chuan_hoa(df["MA_MAY"])

    # Loại toàn bộ MA_MAY cũ có chứa đoạn ".44194" ở bất kỳ vị trí nào.
    bo_qua_dinh_dang = ma_may_anh_xa.str.contains(r"\.44194", regex=True)
    df = df.loc[~bo_qua_dinh_dang].copy()
    ma_may_anh_xa = ma_may_anh_xa.loc[df.index]

    co_gia_tri_hop_le = ma_may_anh_xa.ne("")
    khop_3176 = ma_may_anh_xa.isin(dm_3176)

    df["Khớp MAMAY3176"] = "Bỏ qua"
    df.loc[co_gia_tri_hop_le & khop_3176, "Khớp MAMAY3176"] = "Có"
    df.loc[co_gia_tri_hop_le & ~khop_3176, "Khớp MAMAY3176"] = "Không"

    df["MA_MAY ánh xạ có nhưng DM_MAY không có"] = ""
    df.loc[co_gia_tri_hop_le & ~khop_3176, "MA_MAY ánh xạ có nhưng DM_MAY không có"] = ma_may_anh_xa

    cot_hien_thi = list(df.columns)
    for cot_moi in ["Khớp MAMAY3176", "MA_MAY ánh xạ có nhưng DM_MAY không có"]:
        if cot_moi not in cot_hien_thi:
            cot_hien_thi.append(cot_moi)
    return df[cot_hien_thi].reset_index(drop=True)


def xuat_excel_ma_may(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="DoiChieu_MaMay", index=False)
        ws = writer.sheets["DoiChieu_MaMay"]

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"TEN_DICH_VU", "TEN_TB"}:
                width = 40
            elif ten_cot in {"MA_DICH_VU", "MA_MAY", "MA_BA", "MAMAY3176", "ID"}:
                width = 18
            elif ten_cot in {"Khớp MAMAY3176", "Khớp MAMAY", "Không tồn tại trong DM"}:
                width = 22
            elif ten_cot == "MA_MAY ánh xạ có nhưng DM_MAY không có":
                width = 34
            else:
                width = 20
            ws.column_dimensions[col[0].column_letter].width = width

    return file_xuat
