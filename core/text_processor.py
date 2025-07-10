"""
Module xử lý và chuẩn hóa text cho địa chỉ
FIXED: Direct removal of administrative abbreviations without conversion
OPTIMIZED: Simplified logic for better accuracy
"""
import pandas as pd
import unicodedata
import re
from functools import lru_cache
from config import REGEX_PATTERNS, VIETTAT_MAP, REMOVE_WORDS


def chuan_hoa(text):
    """
    Chuẩn hóa text địa chỉ - FIXED với logic xóa trực tiếp
    
    Args:
        text: Text cần chuẩn hóa
        
    Returns:
        str: Text đã được chuẩn hóa
    """
    if pd.isna(text):
        return ''

    text = str(text).strip()
    
    # Bước 1: Xử lý ký tự đặc biệt trước khi lowercase
    text = text.replace('–', '-').replace('—', '-')  # Normalize dashes
    text = re.sub(r'[""''„"]', '"', text)  # Normalize quotes
    
    # Bước 2: Lowercase
    text = text.lower()
    
    # Bước 3: FIXED - Xóa trực tiếp các viết tắt hành chính (TRƯỚC khi áp dụng VIETTAT_MAP)
    for admin_pattern in REGEX_PATTERNS['admin_removals']:
        text = admin_pattern.sub('', text)
    
    # Bước 4: Áp dụng mapping viết tắt cho địa danh cụ thể (sau khi đã xóa viết tắt hành chính)
    for pattern, replacement in VIETTAT_MAP.items():
        text = re.sub(pattern, replacement, text)
    
    # Bước 5: Xử lý số thứ tự trong tên
    text = re.sub(r'\b(\d+)(st|nd|rd|th)\b', r'\1', text)
    
    # Bước 6: Chuẩn hóa Unicode
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(c for c in text if not unicodedata.combining(c))
    
    # Bước 7: Xử lý số (loại bỏ số 0 đầu)
    text = REGEX_PATTERNS['remove_zeros'].sub(r'\1', text)
    
    # Bước 8: Chuẩn hóa dấu và khoảng trắng
    text = text.replace('-', ' ')
    text = REGEX_PATTERNS['special_chars'].sub(' ', text)
    text = REGEX_PATTERNS['whitespace'].sub(' ', text)
    
    # Bước 9: Làm sạch cuối cùng
    text = text.strip()
    
    # Bước 10: Loại bỏ các từ đơn lẻ không có nghĩa (cuối cùng)
    words = text.split()
    meaningful_words = []
    for word in words:
        if (word.isdigit() or 
            len(word) >= 2 or 
            word not in REMOVE_WORDS):
            meaningful_words.append(word)
    
    return ' '.join(meaningful_words)


def chuan_hoa_dia_chi_chi_tiet(address_text):
    """
    Chuẩn hóa địa chỉ chi tiết (bao gồm số nhà, đường, hẻm)
    
    Args:
        address_text: Địa chỉ chi tiết cần chuẩn hóa
        
    Returns:
        dict: Thông tin địa chỉ đã tách và chuẩn hóa
    """
    if pd.isna(address_text) or not address_text:
        return {
            'so_nha': '',
            'hem_ngo': '',
            'duong': '',
            'ap_thon': '',
            'dia_chi_chuan': ''
        }
    
    result = {
        'so_nha': '',
        'hem_ngo': '',
        'duong': '',
        'ap_thon': '',
        'dia_chi_chuan': ''
    }
    
    # Chuẩn hóa text trước
    normalized_text = chuan_hoa(address_text)
    
    # Extract số nhà
    so_nha_match = REGEX_PATTERNS['so_nha_pattern'].search(normalized_text)
    if so_nha_match:
        result['so_nha'] = so_nha_match.group(2)
    
    # Extract hẻm/ngõ
    hem_ngo_match = REGEX_PATTERNS['hem_ngo_pattern'].search(normalized_text)
    if hem_ngo_match:
        result['hem_ngo'] = f"{hem_ngo_match.group(1)} {hem_ngo_match.group(2)}"
    
    # Extract đường
    duong_match = REGEX_PATTERNS['duong_pattern'].search(normalized_text)
    if duong_match:
        result['duong'] = duong_match.group(2).strip()
    
    # Extract ấp/thôn
    result['ap_thon'] = parse_ap_from_address(address_text)
    
    # Địa chỉ đã chuẩn hóa
    result['dia_chi_chuan'] = normalized_text
    
    return result

def tach_chu_so(text):
    """
    Tách phần chữ và số từ text
    
    Args:
        text: Text cần tách
        
    Returns:
        tuple: (phần chữ, phần số)
    """
    match = re.match(r'^([a-z\s]+?)(?:\s+(\d+))?$', text)
    if match:
        chu = match.group(1).strip()
        so_raw = match.group(2)
        so = str(int(so_raw)) if so_raw and so_raw.isdigit() else ''
        return chu, so
    return text.strip(), ''


@lru_cache(maxsize=10000)
def tach_phanchinh(text):
    """
    Tách phần chính của địa chỉ (loại bỏ các từ đơn vị hành chính)
    
    Args:
        text: Text cần xử lý
        
    Returns:
        str: Text đã loại bỏ đơn vị hành chính
    """
    if not text:
        return ''
    text = text.lower()
    parts = text.strip().split()
    return ' '.join([w for w in parts if w not in REMOVE_WORDS])


def parse_ap_from_address(address_text):
    """
    Parse ấp information from địa chỉ chi tiết
    
    Args:
        address_text: Địa chỉ chi tiết string
        
    Returns:
        str: Ấp name found, or empty string if not found
    """
    if pd.isna(address_text) or not address_text:
        return ''
    
    address_text = str(address_text).strip()
    
    # Find all matches for ấp patterns
    matches = list(REGEX_PATTERNS['ap_pattern'].finditer(address_text))
    
    if not matches:
        return ''
    
    # Take the last match (phía sau từ/cụm từ chỉ điểm đơn vị hành chính)
    last_match = matches[-1]
    keyword = last_match.group(1).lower()
    value = last_match.group(2)
    
    # Normalize the result
    normalized_ap = normalize_ap_value(keyword, value)
    
    return normalized_ap


def normalize_ap_value(keyword, value):
    """
    Normalize ấp value based on keyword
    
    Args:
        keyword: The keyword (ấp, khu phố, KP, etc.)
        value: The value following the keyword
        
    Returns:
        str: Normalized ấp value
    """
    # Normalize numbers: "01" -> "1"
    if value.isdigit():
        value = str(int(value))
    
    # Standardize different keywords
    keyword_lower = keyword.lower()
    if keyword_lower in ['kp', 'khu phố']:
        return f"khu pho {value}"
    elif keyword_lower in ['ấp', 'ap']:
        return f"ap {value}"
    elif keyword_lower in ['thôn', 'thon']:
        return f"thon {value}"
    elif keyword_lower in ['khóm', 'khom']:
        return f"khom {value}"
    else:
        return f"{keyword_lower} {value}"


def find_ap_column(df):
    """
    Find ấp column in DataFrame using fuzzy matching
    
    Args:
        df: DataFrame to search in
        
    Returns:
        str or None: Column name for ấp, or None if not found
    """
    from thefuzz import fuzz
    
    ap_keywords = ['ấp', 'ap', 'thôn', 'thon', 'thôn/ấp', 'thon/ap', 'khu phố', 'khu pho']
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for keyword in ap_keywords:
            if fuzz.partial_ratio(keyword, col_lower) >= 80:
                return col
    
    return None


def find_address_column(df):
    """
    Find địa chỉ column in DataFrame using fuzzy matching
    
    Args:
        df: DataFrame to search in
        
    Returns:
        str or None: Column name for địa chỉ, or None if not found
    """
    from thefuzz import fuzz
    
    address_keywords = ['địa chỉ', 'dia chi', 'địa chỉ chi tiết', 'dia chi chi tiet']
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for keyword in address_keywords:
            if fuzz.partial_ratio(keyword, col_lower) >= 80:
                return col
    
    return None


def validate_input(xa, huyen, tinh):
    """
    Kiểm tra và cảnh báo về dữ liệu đầu vào
    
    Args:
        xa, huyen, tinh: Các thành phần địa chỉ
        
    Returns:
        list: Danh sách cảnh báo
    """
    warnings = []
    
    # Kiểm tra độ dài bất thường
    if xa and len(xa) > 50:
        warnings.append(f"Tên xã quá dài: {xa}")
    if huyen and len(huyen) > 50:
        warnings.append(f"Tên huyện quá dài: {huyen}")
    if tinh and len(tinh) > 50:
        warnings.append(f"Tên tỉnh quá dài: {tinh}")
    
    # Kiểm tra ký tự lạ
    special_chars = re.compile(r'[^a-zA-Z0-9\s\-áàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđ]')
    if xa and special_chars.search(xa):
        warnings.append(f"Xã chứa ký tự đặc biệt: {xa}")
    
    return warnings