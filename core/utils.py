# core/utils.py
import pandas as pd
import os

def chuyen_doi_ngay_7980(date_val):
    if pd.isna(date_val): return ""
    date_str = str(date_val).strip()
    if date_str.endswith('.0'): date_str = date_str[:-2]
    if len(date_str) == 12 and date_str.isdigit():
        return f"{date_str[6:8]}/{date_str[4:6]}/{date_str[0:4]} {date_str[8:10]}:{date_str[10:12]}"
    return date_str

def lam_sach_chuoi(val):
    if pd.isna(val) or val == 'nan': return ""
    val_str = str(val).strip()
    if val_str.endswith('.0') and val_str.replace('.', '', 1).isdigit(): return val_str[:-2]
    return val_str

def doc_file_thong_minh(file_path):
    ten_file = os.path.basename(file_path)
    try:
        # 1. THỬ ĐỌC NHƯ EXCEL CHUẨN (Dành cho file thật)
        try:
            if file_path.endswith('.xlsx'):
                return pd.read_excel(file_path, engine='openpyxl')
            elif file_path.endswith('.xls'):
                return pd.read_excel(file_path, engine='xlrd')
            elif file_path.endswith('.csv'):
                return pd.read_csv(file_path)
            else:
                return pd.read_excel(file_path)
        except Exception:
            pass # Nếu lỗi, bỏ qua và thử các cách "mẹo" bên dưới

        # 2. XỬ LÝ "FAKE EXCEL" DẠNG HTML (Đặc sản của phần mềm y tế)
        # Bản chất là file HTML chứa thẻ <table> nhưng bị ép đuôi .xls
        try:
            dfs = pd.read_html(file_path, encoding='utf-8')
            if dfs and len(dfs) > 0:
                # Thường dữ liệu thật nằm ở bảng đầu tiên hoặc bảng to nhất
                # Ở đây ta lấy bảng đầu tiên
                return dfs[0]
        except Exception:
            pass

        # 3. XỬ LÝ "FAKE EXCEL" DẠNG TEXT PHÂN CÁCH BẰNG TAB
        try:
            return pd.read_csv(file_path, sep='\t', encoding='utf-16')
        except Exception:
            pass
            
        try:
            return pd.read_csv(file_path, sep='\t', encoding='utf-8')
        except Exception:
            pass

        # Nếu thử hết mọi cách vẫn không được
        raise ValueError("Định dạng file không được hỗ trợ hoặc file bị hỏng cấu trúc.")

    except Exception as e:
        raise Exception(f"Lỗi khi đọc file {ten_file}. Chi tiết: {str(e)}")