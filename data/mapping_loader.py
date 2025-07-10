"""
Module load và quản lý dữ liệu mapping
UPDATED: Support new sheet2 structure with ấp columns
FIXED: Store both original and normalized data to preserve Vietnamese characters
"""
import pandas as pd
import os
from core.text_processor import chuan_hoa
from core.fuzzy_matcher import fuzzy_matcher
from utils.helpers import get_mapping_file_path


class MappingLoader:
    """Class quản lý việc load và cache dữ liệu mapping - UPDATED & FIXED"""
    
    def __init__(self):
        self.mapping_sheet1 = []
        self.mapping_sheet2 = []
        self.mapping_sheet1_original = []  # NEW: Store original data
        self.mapping_sheet2_original = []  # NEW: Store original data
        self.is_loaded = False
        self.mapping_file_path = get_mapping_file_path()
    
    def load_mapping(self):
        """
        Load dữ liệu mapping từ file Excel - FIXED to preserve original data
        
        Returns:
            bool: True nếu load thành công
            
        Raises:
            Exception: Nếu có lỗi khi load file
        """
        try:
            if not os.path.exists(self.mapping_file_path):
                raise FileNotFoundError(f"Không tìm thấy file mapping.xlsx tại: {self.mapping_file_path}")
            
            # Đọc dữ liệu từ 2 sheet
            df1 = pd.read_excel(self.mapping_file_path, sheet_name=0)
            df2 = pd.read_excel(self.mapping_file_path, sheet_name=1)
            
            # Validate columns for sheet1 (unchanged)
            required_columns_sheet1 = ['xacu', 'huyencu', 'tinhcu', 'xamoi', 'tinhmoi']
            missing_cols = [col for col in required_columns_sheet1 if col not in df1.columns]
            if missing_cols:
                raise ValueError(f"Sheet 1 thiếu các cột: {missing_cols}")
            
            # Validate columns for sheet2 (NEW structure)
            required_columns_sheet2 = ['apcu', 'xacu', 'huyencu', 'tinhcu', 'apmoi', 'xamoi', 'tinhmoi']
            missing_cols = [col for col in required_columns_sheet2 if col not in df2.columns]
            if missing_cols:
                raise ValueError(f"Sheet 2 thiếu các cột: {missing_cols}")
            
            # FIXED: Store both original and normalized data
            self._process_mapping_data(df1, df2)
            
            # Load vào fuzzy matcher with both original and normalized data
            fuzzy_matcher.load_mapping_data(
                self.mapping_sheet1, self.mapping_sheet2,
                self.mapping_sheet1_original, self.mapping_sheet2_original
            )
            
            self.is_loaded = True
            return True
            
        except Exception as e:
            raise Exception(f"Lỗi khi load file mapping: {str(e)}")
    
    def _process_mapping_data(self, df1, df2):
        """Process mapping data - FIXED: Store both original and normalized versions"""
        # Process Sheet1
        # Store original data (with Vietnamese characters)
        self.mapping_sheet1_original = list(zip(
            df1['xacu'].fillna('').astype(str),
            df1['huyencu'].fillna('').astype(str), 
            df1['tinhcu'].fillna('').astype(str),
            df1['xamoi'].fillna('').astype(str), 
            df1['tinhmoi'].fillna('').astype(str)
        ))
        
        # Store normalized data (for matching)
        for col in ['xacu', 'huyencu', 'tinhcu', 'xamoi', 'tinhmoi']:
            df1[f'{col}_chuan'] = df1[col].apply(chuan_hoa)
        
        self.mapping_sheet1 = list(zip(
            df1['xacu_chuan'], df1['huyencu_chuan'], df1['tinhcu_chuan'],
            df1['xamoi_chuan'], df1['tinhmoi_chuan']
        ))
        
        # Process Sheet2
        # Store original data (with Vietnamese characters)
        self.mapping_sheet2_original = list(zip(
            df2['apcu'].fillna('').astype(str),
            df2['xacu'].fillna('').astype(str),
            df2['huyencu'].fillna('').astype(str),
            df2['tinhcu'].fillna('').astype(str),
            df2['apmoi'].fillna('').astype(str),
            df2['xamoi'].fillna('').astype(str),
            df2['tinhmoi'].fillna('').astype(str)
        ))
        
        # Store normalized data (for matching)
        for col in ['apcu', 'xacu', 'huyencu', 'tinhcu', 'apmoi', 'xamoi', 'tinhmoi']:
            df2[f'{col}_chuan'] = df2[col].apply(chuan_hoa)
        
        self.mapping_sheet2 = list(zip(
            df2['apcu_chuan'], df2['xacu_chuan'], df2['huyencu_chuan'], df2['tinhcu_chuan'],
            df2['apmoi_chuan'], df2['xamoi_chuan'], df2['tinhmoi_chuan']
        ))
    
    def get_mapping_stats(self):
        """
        Lấy thống kê về dữ liệu mapping
        
        Returns:
            dict: Thống kê mapping
        """
        if not self.is_loaded:
            return None
        
        return {
            'sheet1_count': len(self.mapping_sheet1),
            'sheet2_count': len(self.mapping_sheet2),
            'total_count': len(self.mapping_sheet1) + len(self.mapping_sheet2),
            'file_path': self.mapping_file_path,
            'file_exists': os.path.exists(self.mapping_file_path)
        }
    
    def reload_mapping(self):
        """Reload dữ liệu mapping"""
        self.is_loaded = False
        return self.load_mapping()
    
    def get_unique_provinces(self):
        """Lấy danh sách tỉnh thành unique - using original data"""
        if not self.is_loaded:
            return []
        
        provinces = set()
        # Use original data to preserve Vietnamese characters
        for mapping in [self.mapping_sheet1_original, self.mapping_sheet2_original]:
            for item in mapping:
                if len(item) == 5:  # Sheet1 format
                    provinces.add(item[2])  # tinhcu
                    provinces.add(item[4])  # tinhmoi
                elif len(item) == 7:  # Sheet2 format
                    provinces.add(item[3])  # tinhcu
                    provinces.add(item[6])  # tinhmoi
        
        return sorted(list(provinces))
    
    def get_mapping_for_province(self, province_name):
        """
        Lấy mapping cho một tỉnh cụ thể
        
        Args:
            province_name: Tên tỉnh
            
        Returns:
            list: Danh sách mapping cho tỉnh đó (original data)
        """
        if not self.is_loaded:
            return []
        
        result = []
        province_normalized = chuan_hoa(province_name)
        
        # Check sheet1 - use normalized for comparison but return original
        for i, item_norm in enumerate(self.mapping_sheet1):
            if (item_norm[2] == province_normalized or   # tinhcu
                item_norm[4] == province_normalized):    # tinhmoi
                result.append(self.mapping_sheet1_original[i])
        
        # Check sheet2 - use normalized for comparison but return original
        for i, item_norm in enumerate(self.mapping_sheet2):
            if (item_norm[3] == province_normalized or   # tinhcu
                item_norm[6] == province_normalized):    # tinhmoi
                result.append(self.mapping_sheet2_original[i])
        
        return result


# Global mapping loader instance
mapping_loader = MappingLoader()


def load_mapping():
    """
    Wrapper function để maintain compatibility với code cũ
    
    Returns:
        bool: True nếu load thành công
    """
    return mapping_loader.load_mapping()


def get_mapping_data():
    """
    Lấy dữ liệu mapping đã load
    
    Returns:
        tuple: (mapping_sheet1, mapping_sheet2, mapping_sheet1_original, mapping_sheet2_original)
    """
    if not mapping_loader.is_loaded:
        load_mapping()
    
    return (mapping_loader.mapping_sheet1, mapping_loader.mapping_sheet2,
            mapping_loader.mapping_sheet1_original, mapping_loader.mapping_sheet2_original)