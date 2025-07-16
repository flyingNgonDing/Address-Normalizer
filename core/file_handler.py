"""
Module xử lý các thao tác file I/O
UPDATED: Enhanced .xls support with proper engine handling
"""
import pandas as pd
import os
from thefuzz import fuzz
from config import SUPPORTED_EXTENSIONS


def read_file(file_path, sheet_name=None):
    """
    Đọc file Excel hoặc CSV - UPDATED with enhanced .xls support
    
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
        if file_ext == '.xlsx':
            # Use openpyxl for .xlsx files
            if sheet_name is not None:
                return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            else:
                return pd.read_excel(file_path, engine='openpyxl')
        elif file_ext == '.xls':
            # Use xlrd for .xls files
            if sheet_name is not None:
                return pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
            else:
                return pd.read_excel(file_path, engine='xlrd')
        else:  # .csv
            return pd.read_csv(file_path, encoding='utf-8')
    except Exception as e:
        # Enhanced error handling with specific suggestions
        error_msg = f"Lỗi đọc file {file_path}: {str(e)}"
        
        if file_ext == '.xls':
            error_msg += f"\n\nGợi ý cho file .xls:"
            error_msg += f"\n- Đảm bảo đã cài đặt xlrd: pip install xlrd==2.0.1"
            error_msg += f"\n- Thử mở file bằng Excel và lưu lại định dạng .xlsx"
            error_msg += f"\n- Kiểm tra file có bị hỏng không"
        
        raise Exception(error_msg)


def get_excel_sheet_names(file_path):
    """
    Lấy danh sách tên sheets trong file Excel - UPDATED with enhanced .xls support
    
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
        # Use appropriate engine based on file extension
        if file_ext == '.xlsx':
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        elif file_ext == '.xls':
            excel_file = pd.ExcelFile(file_path, engine='xlrd')
        
        return excel_file.sheet_names
    except Exception as e:
        error_msg = f"Lỗi đọc danh sách sheets từ {file_path}: {str(e)}"
        
        if file_ext == '.xls':
            error_msg += f"\n\nGợi ý cho file .xls:"
            error_msg += f"\n- Đảm bảo đã cài đặt xlrd: pip install xlrd==2.0.1"
            error_msg += f"\n- Kiểm tra file có bị corrupt không"
        
        raise Exception(error_msg)


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
        # Always save as .xlsx for better compatibility
        df.to_excel(file_path, index=False, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Lỗi lưu file {file_path}: {str(e)}")


def save_multiple_sheets(sheet_data_dict, file_path):
    """
    Lưu multiple sheets vào file Excel
    
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


def test_excel_engines():
    """
    Test function to check available Excel engines
    
    Returns:
        dict: Status of each engine
    """
    engines = {}
    
    # Test openpyxl
    try:
        import openpyxl
        engines['openpyxl'] = f"✅ Available (version: {openpyxl.__version__})"
    except ImportError:
        engines['openpyxl'] = "❌ Not available"
    
    # Test xlrd
    try:
        import xlrd
        engines['xlrd'] = f"✅ Available (version: {xlrd.__version__})"
    except ImportError:
        engines['xlrd'] = "❌ Not available"
    
    # Test xlsxwriter
    try:
        import xlsxwriter
        engines['xlsxwriter'] = f"✅ Available (version: {xlsxwriter.__version__})"
    except ImportError:
        engines['xlsxwriter'] = "❌ Not available"
    
    return engines


def diagnose_file_issue(file_path):
    """
    Diagnose issues with Excel file reading
    
    Args:
        file_path: Path to the problematic file
        
    Returns:
        dict: Diagnostic information
    """
    diagnosis = {
        'file_exists': os.path.exists(file_path),
        'file_size': 0,
        'file_extension': '',
        'engines': test_excel_engines(),
        'pandas_version': pd.__version__,
        'suggestions': []
    }
    
    if diagnosis['file_exists']:
        diagnosis['file_size'] = os.path.getsize(file_path)
        diagnosis['file_extension'] = os.path.splitext(file_path.lower())[1]
        
        # Try to detect file corruption
        try:
            with open(file_path, 'rb') as f:
                header = f.read(8)
                if diagnosis['file_extension'] == '.xls':
                    # Check for XLS signature
                    if header[:2] != b'\xd0\xcf':
                        diagnosis['suggestions'].append("File may be corrupted or not a valid .xls file")
                elif diagnosis['file_extension'] == '.xlsx':
                    # Check for ZIP signature (XLSX is a ZIP file)
                    if header[:2] != b'PK':
                        diagnosis['suggestions'].append("File may be corrupted or not a valid .xlsx file")
        except Exception as e:
            diagnosis['suggestions'].append(f"Cannot read file header: {e}")
    
    # Add engine-specific suggestions
    if diagnosis['file_extension'] == '.xls':
        if 'xlrd' in diagnosis['engines'] and '❌' in diagnosis['engines']['xlrd']:
            diagnosis['suggestions'].append("Install xlrd for .xls support: pip install xlrd==2.0.1")
    
    if diagnosis['file_extension'] == '.xlsx':
        if 'openpyxl' in diagnosis['engines'] and '❌' in diagnosis['engines']['openpyxl']:
            diagnosis['suggestions'].append("Install openpyxl for .xlsx support: pip install openpyxl")
    
    return diagnosis
