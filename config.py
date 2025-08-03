"""
Configuration và constants cho chương trình chuẩn hóa địa chỉ
WINDOWS VERSION - Optimized for Windows 8/10/11
FIXED: Direct removal of administrative abbreviations without conversion
"""
import os
import sys
import re
from pathlib import Path

# Đường dẫn ứng dụng - WINDOWS VERSION
if getattr(sys, 'frozen', False):
    # Running in a bundle (PyInstaller)
    APPLICATION_PATH = os.path.dirname(sys.executable)
    # Alternative path for some PyInstaller configurations
    if hasattr(sys, '_MEIPASS'):
        APPLICATION_PATH = sys._MEIPASS
else:
    # Running in development
    APPLICATION_PATH = os.path.dirname(os.path.abspath(__file__))

# Ensure APPLICATION_PATH uses Windows path format
APPLICATION_PATH = os.path.normpath(APPLICATION_PATH)

# Pre-compiled regex patterns để tối ưu hiệu suất - FIXED for direct removal
REGEX_PATTERNS = {
    'remove_zeros': re.compile(r'\b0+(\d+)\b'),
    'whitespace': re.compile(r'\s+'),
    'special_chars': re.compile(r'[,.]'),
    'word_boundary': re.compile(r'\b'),
    
    # FIXED: Remove ONLY abbreviations (with dots), keep full words
    'admin_removals': [
        # Thành phố - ONLY abbreviations with dots
        re.compile(r'\btp[.]\s*', re.IGNORECASE),       # TP. (with dot)
        
        # Thị xã - ONLY abbreviations with dots
        re.compile(r'\btx[.]\s*', re.IGNORECASE),       # TX. (with dot)
        
        # Thị trấn - ONLY abbreviations with dots
        re.compile(r'\btt[.]\s*', re.IGNORECASE),       # TT. (with dot)
        
        # Quận - ONLY abbreviations with dots
        re.compile(r'\bq[.]\s*', re.IGNORECASE),        # Q. (with dot)
        
        # Phường - ONLY abbreviations with dots
        re.compile(r'\bp[.]\s*', re.IGNORECASE),        # P. (with dot)
        
        # Xã - ONLY abbreviations with dots
        re.compile(r'\bx[.]\s*', re.IGNORECASE),        # X. (with dot)
        
        # Huyện - ONLY abbreviations with dots
        re.compile(r'\bh[.]\s*', re.IGNORECASE),        # H. (with dot and space only)
        
        # Additional administrative units - ONLY with dots
        re.compile(r'\bkp[.]\s*', re.IGNORECASE),       # KP. (with dot)
        re.compile(r'\btdp[.]\s*', re.IGNORECASE),      # TDP. (with dot)
        re.compile(r'\bkv[.]\s*', re.IGNORECASE),       # KV. (with dot)
    ],
    
    # Patterns for ấp parsing - unchanged
    'ap_pattern': re.compile(
        r'\b(ấp|ap|thôn|thon|khu phố|khu pho|kp|khóm|khom|tổ|to|bản|ban|'
        r'tổ dân phố|to dan pho|tdp|khu vực|khu vuc|kv)\s+([a-zA-Z0-9]+)',
        re.IGNORECASE
    ),
    'number_normalize': re.compile(r'^0+(\d+)$'),
    
    # Patterns for detailed address parsing - unchanged
    'so_nha_pattern': re.compile(r'\b(số|so|s\.)\s*(\d+[a-z]*(?:/\d+[a-z]*)*)', re.IGNORECASE),
    'duong_pattern': re.compile(r'\b(đường|duong|đ\.|d\.)\s+([^,]+)', re.IGNORECASE),
    'hem_ngo_pattern': re.compile(r'\b(hẻm|hem|h\.|ngõ|ngo|n\.)\s*(\d+[a-z]*(?:/\d+[a-z]*)*)', re.IGNORECASE),
}

# REMOVED: DONVI_MAP - No longer needed since we do direct removal
# REMOVED: LOAI_BO_WORDS - Simplified approach

# Mapping từ đầy đủ sang chuẩn hóa - NEW: Handle full words
FULLWORD_MAP = {
    # Administrative units - full words to normalized form
    r'\bphường\b': 'phuong',
    r'\bxã\b': 'xa', 
    r'\bthị trấn\b': 'thi tran',
    r'\bthị xã\b': 'thi xa',
    r'\bthành phố\b': 'thanh pho',
    r'\bquận\b': 'quan',
    r'\bhuyện\b': 'huyen',
    r'\btỉnh\b': 'tinh',
    
    # Additional units
    r'\bấp\b': 'ap',
    r'\bthôn\b': 'thon',
    r'\bbản\b': 'ban',
    r'\bkhóm\b': 'khom',
    r'\btổ\b': 'to',
    r'\bkhu phố\b': 'khu pho',
    r'\btổ dân phố\b': 'to dan pho',
    r'\bkhu vực\b': 'khu vuc',
}
VIETTAT_MAP = {
    # TP.HCM variations - ENHANCED
    r'\btp[.]?\s*h[.]?\s*c[.]?\s*m[.]?\b': 'thanh pho ho chi minh',
    r'\btphcm\b': 'thanh pho ho chi minh',
    r'\bhcm\b': 'thanh pho ho chi minh',
    r'\btp[.]?\s*hcm\b': 'thanh pho ho chi minh',
    r'\bh[.]?\s*c[.]?\s*m[.]?\b': 'thanh pho ho chi minh',
    r'\bsai gon\b': 'thanh pho ho chi minh',
    r'\bsaigon\b': 'thanh pho ho chi minh',
    r'\bsg\b': 'thanh pho ho chi minh',
    
    # Bà Rịa - Vũng Tàu
    r'\bbr[ -]?vt\b': 'ba ria vung tau',
    r'\bb[.]?\s*ria[ -]?v[.]?\s*tau\b': 'ba ria vung tau',
    r'\bba ria[ -]?vung tau\b': 'ba ria vung tau',
    
    # Các tỉnh thành phổ biến khác
    r'\bhn\b': 'ha noi',
    r'\bdn\b': 'da nang',
    r'\bvt\b': 'vung tau',
    r'\bbd\b': 'binh duong',
    r'\bbduong\b': 'binh duong',
    r'\bla\b': 'long an',
    r'\btg\b': 'tien giang',
    r'\bct\b': 'can tho',
    r'\bag\b': 'an giang',
    r'\bkg\b': 'kien giang',
    r'\bcm\b': 'ca mau',
    r'\bbl\b': 'bac lieu',
    r'\btv\b': 'tra vinh',
    r'\bst\b': 'soc trang',
    r'\bdt\b': 'dong thap',
    r'\bvl\b': 'vinh long',
    r'\bht\b': 'hau giang',
    r'\bbn\b': 'ben tre',
    
    # Xử lý số thứ tự
    r'\b(\d+)st\b': r'\1',
    r'\b(\d+)nd\b': r'\1', 
    r'\b(\d+)rd\b': r'\1',
    r'\b(\d+)th\b': r'\1',
}

# SIMPLIFIED: Only essential words to remove at the end
REMOVE_WORDS = [
    'xa', 'phuong', 'huyen', 'quan', 'tinh', 'thi tran', 'ap', 'thon', 'ban', 'khom', 'to'
]

# UI Configuration - Windows optimized
DEFAULT_GEOMETRY = "750x350"
EXPANDED_GEOMETRY = "750x530"

# Colors - Windows friendly với contrast cao
COLORS = {
    'bg_primary': '#ffffff',       # White background for better readability
    'bg_drag': '#e3f2fd',          # Light blue for drag indication
    'text_primary': '#1a1a1a',     # Dark text for high contrast
    'text_secondary': '#4a4a4a',   # Medium gray text
    'accent': '#1976d2',           # Blue accent
    'main_button': '#0d47a1',      # Dark blue for main button
    'success': '#2e7d32',          # Dark green
    'danger': '#d32f2f',           # Dark red
    'warning': '#f57c00',          # Orange
    'button_bg': '#f5f5f5',        # Light gray for buttons
    'button_border': '#cccccc',    # Border color
}

# File extensions hỗ trợ
SUPPORTED_EXTENSIONS = ('.xlsx', '.xls', '.csv')

# Processing configuration - Windows optimized
CHUNK_SIZE = 500  # Smaller chunks for Windows
MAX_WORKERS = min(4, os.cpu_count() or 1)  # Limit workers on Windows

# Thresholds cho fuzzy matching
FUZZY_THRESHOLDS = {
    'xa_min': 85,
    'huyen_min': 70,
    'tinh_min': 60,
    'total_min': 75,
    'xa_moi_min': 85,
    'tinh_moi_min': 70,
    # Thresholds for ấp
    'ap_min': 75,      # For text-based ấp names
    'ap_number_exact': True,  # Numbers must match exactly (after normalization)
}

# Animation settings - Reduced for Windows performance
ANIMATION_SETTINGS = {
    'steps': 10,        # Fewer steps for smoother performance
    'delay': 30,        # Slightly longer delay
}

# Windows specific settings
WINDOWS_SETTINGS = {
    'use_native_dialogs': True,
    'disable_dpi_awareness': False,
    'font_scaling': 1.0,
}