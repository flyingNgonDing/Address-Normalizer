"""
Module xử lý các thao tác file I/O
UPDATED: Added .xls support and sheet selection functionality
"""
import pandas as pd
import os
from thefuzz import fuzz
from config import SUPPORTED_EXTENSIONS


def read_file(file_path, sheet_name=None):
    """
    Đọc file Excel hoặc CSV - UPDATED with .xls support and sheet selection
    
    Args:
        file_path: Đường dẫn file
        sheet_name: Tên sheet cần đọc (None để đọc sheet đầu tiên)
        
    Returns:
        pd.DataFrame: Dữ liệu đã đọc
        
    Raises:
        Exception: Nếu không thể đọc file
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File không tồn tại: {file_path}")
    
    file_ext = os.path.splitext(file_path.lower())[1]
    
    if file_ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Định dạng file không được hỗ trợ: {file_ext}")
    
    try:
        if file_ext in ['.xlsx', '.xls', 
                       .csv']:
            # Support both .xlsx and .xls files
            if sheet_name is not None:
                return pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                return pd.read_excel(file_path)
        else:  # .csv
            return pd.read_csv(file_path, encoding='utf-8')
    except Exception as e:
        raise Exception(f"Lỗi đọc file {file_path}: {str(e)}")


def get_excel_sheet_names(file_path):
    """
    NEW: Lấy danh sách tên sheets trong file Excel
    
    Args:
        file_path: Đường dẫn file Excel
        
    Returns:
        list: Danh sách tên sheets
        
    Raises:
        Exception: Nếu không thể đọc file hoặc không phải file Excel
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File không tồn tại: {file_path}")
    
    file_ext = os.path.splitext(file_path.lower())[1]
    
    if file_ext not in ['.xlsx', '.xls']:
        raise ValueError(f"File không phải Excel: {file_ext}")
    
    try:
        # Use ExcelFile to get sheet names without loading data
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        raise Exception(f"Lỗi đọc danh sách sheets từ {file_path}: {str(e)}")


def save_file(df, file_path):
    """
    Lưu DataFrame thành file Excel
    
    Args:
        df: DataFrame cần lưu
        file_path: Đường dẫn file đích
        
    Raises:
        Exception: Nếu không thể lưu file
    """
    try:
        df.to_excel(file_path, index=False)
    except Exception as e:
        raise Exception(f"Lỗi lưu file {file_path}: {str(e)}")


def save_multiple_sheets(sheet_data_dict, file_path):
    """
    NEW: Lưu multiple sheets vào file Excel
    
    Args:
        sheet_data_dict: Dictionary {sheet_name: DataFrame}
        file_path: Đường dẫn file đích
        
    Raises:
        Exception: Nếu không thể lưu file
    """
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in sheet_data_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        raise Exception(f"Lỗi lưu file {file_path}: {str(e)}")


def tim_cot(df, ten_goc):
    """
    Tìm cột trong DataFrame dựa trên tên gần đúng
    
    Args:
        df: DataFrame
        ten_goc: Tên cột cần tìm
        
    Returns:
        str or None: Tên cột tìm được hoặc None
    """
    for col in df.columns:
        if fuzz.partial_ratio(ten_goc.lower(), col.lower()) >= 80:
            return col
    return None


def validate_file_format(file_path):
    """
    Kiểm tra định dạng file có hợp lệ không
    
    Args:
        file_path: Đường dẫn file
        
    Returns:
        bool: True nếu hợp lệ
    """
    if not file_path:
        return False
    
    file_ext = os.path.splitext(file_path.lower())[1]
    return file_ext in SUPPORTED_EXTENSIONS


def get_file_info(file_path):
    """
    Lấy thông tin cơ bản về file
    
    Args:
        file_path: Đường dẫn file
        
    Returns:
        dict: Thông tin file
    """
    if not os.path.exists(file_path):
        return None
    
    stat = os.stat(file_path)
    return {
        'name': os.path.basename(file_path),
        'size': stat.st_size,
        'extension': os.path.splitext(file_path.lower())[1],
        'path': file_path
    }


def check_required_columns(df):
    """
    Kiểm tra file có đủ các cột bắt buộc không
    
    Args:
        df: DataFrame cần kiểm tra
        
    Returns:
        dict: {'valid': bool, 'xa_col': str, 'huyen_col': str, 'tinh_col': str, 'missing': list}
    """
    required_columns = {'xã': 'xa', 'huyện': 'huyen', 'tỉnh': 'tinh'}
    found_columns = {}
    missing = []
    
    for display_name, key in required_columns.items():
        col = tim_cot(df, display_name)
        if col:
            found_columns[f'{key}_col'] = col
        else:
            missing.append(display_name)
    
    return {
        'valid': len(missing) == 0,
        'missing': missing,
        **found_columns
    }
