"""
Module xử lý fuzzy matching cho địa chỉ
UPDATED: Support new sheet2 structure with ấp matching
FIXED: Return original data with Vietnamese characters instead of normalized data
FIXED: Syntax errors and logic issues
"""
from thefuzz import fuzz
from core.text_processor import tach_chu_so, tach_phanchinh, normalize_ap_value, parse_ap_from_address
from config import FUZZY_THRESHOLDS


class FuzzyMatcher:
    """Class xử lý fuzzy matching với cache - FIXED to preserve Vietnamese characters"""
    
    def __init__(self):
        self.cache = {}
        self.new_address_cache = {}
        self.sheet2_cache = {}
        self.mapping_sheet1 = []           # Normalized data for matching
        self.mapping_sheet2 = []           # Normalized data for matching
        self.mapping_sheet1_original = []  # NEW: Original data for output
        self.mapping_sheet2_original = []  # NEW: Original data for output
    
    def load_mapping_data(self, mapping_sheet1, mapping_sheet2, mapping_sheet1_original=None, mapping_sheet2_original=None):
        """
        Load dữ liệu mapping và tạo cache - UPDATED to handle original data
        
        Args:
            mapping_sheet1: Dữ liệu mapping từ sheet1 (5 elements, normalized)
            mapping_sheet2: Dữ liệu mapping từ sheet2 (7 elements, normalized)
            mapping_sheet1_original: Original data từ sheet1 (5 elements, with Vietnamese chars)
            mapping_sheet2_original: Original data từ sheet2 (7 elements, with Vietnamese chars)
        """
        self.mapping_sheet1 = mapping_sheet1
        self.mapping_sheet2 = mapping_sheet2
        
        # FIXED: Store original data for output
        self.mapping_sheet1_original = mapping_sheet1_original or mapping_sheet1
        self.mapping_sheet2_original = mapping_sheet2_original or mapping_sheet2
        
        self._build_cache()
    
    def _build_cache(self):
        """Tạo cache để tối ưu hiệu suất - using normalized data for matching"""
        self.cache = {}
        self.new_address_cache = {}
        self.sheet2_cache = {}
        
        # Build cache for sheet1 (using normalized data)
        for i, item in enumerate(self.mapping_sheet1):
            xacu, huyencu, tinhcu, xamoi, tinhmoi = item
            
            # Cache cho địa chỉ cũ (xã cũ + huyện + tỉnh cũ)
            key = (tach_phanchinh(xacu), tach_phanchinh(huyencu), tach_phanchinh(tinhcu))
            if key not in self.cache:
                self.cache[key] = []
            self.cache[key].append((i, 'sheet1'))  # Store index instead of item
            
            # Cache cho địa chỉ mới (xã mới + tỉnh mới)
            new_key = (tach_phanchinh(xamoi), tach_phanchinh(tinhmoi))
            if new_key not in self.new_address_cache:
                self.new_address_cache[new_key] = []
            self.new_address_cache[new_key].append((i, 'sheet1'))  # Store index instead of item
        
        # Build cache for sheet2 (using normalized data)
        for i, item in enumerate(self.mapping_sheet2):
            apcu, xacu, huyencu, tinhcu, apmoi, xamoi, tinhmoi = item
            
            # Cache cho địa chỉ cũ với ấp (ấp cũ + xã cũ + huyện + tỉnh cũ)
            key = (
                tach_phanchinh(apcu), 
                tach_phanchinh(xacu), 
                tach_phanchinh(huyencu), 
                tach_phanchinh(tinhcu)
            )
            if key not in self.sheet2_cache:
                self.sheet2_cache[key] = []
            self.sheet2_cache[key].append(i)  # Store index instead of item
    
    def match_row(self, xa, huyen, tinh, ap=None, address_detail=None):
        """
        Thực hiện fuzzy matching cho một dòng dữ liệu - FIXED to return original data
        
        Args:
            xa, huyen, tinh: Thông tin địa chỉ cần match
            ap: Thông tin ấp (optional)
            address_detail: Địa chỉ chi tiết để parse ấp (optional)
            
        Returns:
            tuple: (xacu, huyencu, tinhcu, xamoi, tinhmoi, lý do) - with original Vietnamese characters
        """
        xa_chu, xa_so = tach_chu_so(tach_phanchinh(xa))
        huyen_chu = tach_phanchinh(huyen)
        tinh_chu = tach_phanchinh(tinh)

        # Determine ấp information
        ap_info = self._get_ap_info(ap, address_detail)

        # XỬ LÝ TRƯỜNG HỢP THIẾU HUYỆN - KIỂM TRA VỚI ĐỊA CHỈ MỚI
        if not huyen_chu.strip():
            return self._match_new_address(xa_chu, xa_so, tinh_chu)
        
        # XỬ LÝ TRƯỜNG HỢP CÓ ẤP - MATCH VỚI SHEET2 TRƯỚC
        if ap_info:
            result = self._match_sheet2_with_ap(xa_chu, xa_so, huyen_chu, tinh_chu, ap_info)
            if result[0] is not None:  # Found match in sheet2
                return result
        
        # XỬ LÝ TRƯỜNG HỢP CÓ ĐẦY ĐỦ XÃ + HUYỆN + TỈNH - MATCH VỚI SHEET1
        return self._match_full_address(xa_chu, xa_so, huyen_chu, tinh_chu, xa, huyen, tinh)
    
    def _get_ap_info(self, ap, address_detail):
        """Get ấp information from ap column or parse from address detail"""
        # Priority 1: Direct ấp column
        if ap and str(ap).strip():
            return tach_phanchinh(str(ap).strip().lower())
        
        # Priority 2: Parse from address detail
        if address_detail:
            parsed_ap = parse_ap_from_address(address_detail)
            if parsed_ap:
                return tach_phanchinh(parsed_ap)
        
        return None
    
    def _match_sheet2_with_ap(self, xa_chu, xa_so, huyen_chu, tinh_chu, ap_info):
        """Match với sheet2 data using ấp information - FIXED to return original data"""
        # Exact match first
        exact_key = (ap_info, xa_chu + ' ' + xa_so if xa_so else xa_chu, huyen_chu, tinh_chu)
        if exact_key in self.sheet2_cache:
            indices = self.sheet2_cache[exact_key]
            if indices:
                # FIXED: Return original data instead of normalized
                original_item = self.mapping_sheet2_original[indices[0]]
                # Convert sheet2 to sheet1 format using original data
                return (original_item[1], original_item[2], original_item[3], original_item[5], original_item[6], '')
        
        # Fuzzy match
        return self._fuzzy_match_sheet2(xa_chu, xa_so, huyen_chu, tinh_chu, ap_info)
    
    def _fuzzy_match_sheet2(self, xa_chu, xa_so, huyen_chu, tinh_chu, ap_info):
        """Fuzzy match với sheet2 data - FIXED to return original data"""
        best_score = 0
        best_match_index = None
        
        for i, item in enumerate(self.mapping_sheet2):
            apcu, xacu, huyencu, tinhcu, apmoi, xamoi, tinhmoi = item
            
            # Parse components (using normalized data for comparison)
            apcu_chu = tach_phanchinh(apcu)
            xacu_chu, xacu_so = tach_chu_so(tach_phanchinh(xacu))
            
            # Filter by number first
            if xa_so and xa_so != xacu_so:
                continue
            
            # Score ấp (highest priority)
            score_ap = self._score_ap_match(ap_info, apcu_chu)
            if score_ap < FUZZY_THRESHOLDS['ap_min']:
                continue
            
            # Score xã
            score_xa = max(
                fuzz.ratio(xa_chu, xacu_chu),
                fuzz.partial_ratio(xa_chu, xacu_chu),
                fuzz.token_sort_ratio(xa_chu, xacu_chu) * 0.9
            )
            if score_xa < FUZZY_THRESHOLDS['xa_min']:
                continue
            
            # Score huyện
            score_huyen = max(
                fuzz.ratio(huyen_chu, tach_phanchinh(huyencu)),
                fuzz.partial_ratio(huyen_chu, tach_phanchinh(huyencu)),
                fuzz.token_sort_ratio(huyen_chu, tach_phanchinh(huyencu)) * 0.9
            )
            if score_huyen < FUZZY_THRESHOLDS['huyen_min']:
                continue
            
            # Score tỉnh
            score_tinh = max(
                fuzz.ratio(tinh_chu, tach_phanchinh(tinhcu)),
                fuzz.partial_ratio(tinh_chu, tach_phanchinh(tinhcu)),
                fuzz.token_sort_ratio(tinh_chu, tach_phanchinh(tinhcu)) * 0.9
            )
            if score_tinh < FUZZY_THRESHOLDS['tinh_min']:
                continue
            
            # Calculate total score with ấp priority
            total_score = (
                0.4 * score_ap +      # Ấp highest priority
                0.3 * score_xa +      # Xã second priority  
                0.2 * score_huyen +   # Huyện third priority
                0.1 * score_tinh      # Tỉnh lowest priority
            )
            
            if total_score > best_score:
                best_score = total_score
                best_match_index = i
        
        if best_match_index is not None and best_score >= FUZZY_THRESHOLDS['total_min']:
            # FIXED: Return original data instead of normalized
            original_item = self.mapping_sheet2_original[best_match_index]
            # Convert sheet2 to sheet1 format using original data
            return (original_item[1], original_item[2], original_item[3], original_item[5], original_item[6], '')
        
        return (None, None, None, None, None, 'xã cấu véo, cần thực hiện thủ công')
    
    def _score_ap_match(self, ap_info, apcu_chu):
        """Score ấp matching with special rules for numbers vs text"""
        if not ap_info or not apcu_chu:
            return 0
        
        # Check if both are number-based (KP, khóm with numbers)
        ap_parts = ap_info.split()
        apcu_parts = apcu_chu.split()
        
        if len(ap_parts) >= 2 and len(apcu_parts) >= 2:
            ap_keyword, ap_value = ap_parts[0], ap_parts[1]
            apcu_keyword, apcu_value = apcu_parts[0], apcu_parts[1]
            
            # For number-based ấp, require exact number match
            if (ap_keyword in ['khu', 'khom'] and apcu_keyword in ['khu', 'khom'] and 
                ap_value.isdigit() and apcu_value.isdigit()):
                if FUZZY_THRESHOLDS['ap_number_exact']:
                    return 100 if ap_value == apcu_value else 0
            
            # For text-based ấp, use fuzzy matching
            if ap_keyword == apcu_keyword:
                return max(
                    fuzz.ratio(ap_value, apcu_value),
                    fuzz.partial_ratio(ap_value, apcu_value)
                )
        
        # Fallback to general fuzzy matching
        return max(
            fuzz.ratio(ap_info, apcu_chu),
            fuzz.partial_ratio(ap_info, apcu_chu),
            fuzz.token_sort_ratio(ap_info, apcu_chu) * 0.9
        )
    
    def _match_new_address(self, xa_chu, xa_so, tinh_chu):
        """Match với địa chỉ mới khi thiếu huyện - FIXED to return original data"""
        new_address_key = (xa_chu + ' ' + xa_so if xa_so else xa_chu, tinh_chu)
        
        # Kiểm tra chính xác với địa chỉ mới
        if new_address_key in self.new_address_cache:
            matches = self.new_address_cache[new_address_key]
            # Ưu tiên sheet1 nếu có
            for index, sheet in matches:
                if sheet == 'sheet1':
                    # FIXED: Return original data
                    original_item = self.mapping_sheet1_original[index]
                    return original_item + ('Địa chỉ mới - giữ nguyên',)
            # Nếu không có trong sheet1, lấy từ sheet2
            if matches:
                index, sheet = matches[0]
                if sheet == 'sheet2':
                    # FIXED: Return original data, convert sheet2 to sheet1 format
                    original_item = self.mapping_sheet2_original[index]
                    return (original_item[1], original_item[2], original_item[3], 
                           original_item[5], original_item[6], 'Địa chỉ mới - giữ nguyên')
        
        # Nếu không match chính xác, thử fuzzy match với địa chỉ mới
        return self._fuzzy_match_new_address(xa_chu, xa_so, tinh_chu)
    
    def _fuzzy_match_new_address(self, xa_chu, xa_so, tinh_chu):
        """Fuzzy match với địa chỉ mới - FIXED to return original data"""
        best_score = 0
        best_match_index = None
        best_sheet = None
        
        # Check sheet1
        for i, item in enumerate(self.mapping_sheet1):
            xacu, huyencu, tinhcu, xamoi, tinhmoi = item
            xa_moi_chu, xa_moi_so = tach_chu_so(tach_phanchinh(xamoi))
            
            # Lọc sơ bộ theo số
            if xa_so and xa_so != xa_moi_so:
                continue
            
            # Tính điểm cho xã mới
            score_xa_moi = max(
                fuzz.ratio(xa_chu, xa_moi_chu),
                fuzz.partial_ratio(xa_chu, xa_moi_chu),
                fuzz.token_sort_ratio(xa_chu, xa_moi_chu) * 0.9
            )
            
            if score_xa_moi < FUZZY_THRESHOLDS['xa_moi_min']:
                continue

            # Tính điểm cho tỉnh mới
            score_tinh_moi = max(
                fuzz.ratio(tinh_chu, tach_phanchinh(tinhmoi)),
                fuzz.partial_ratio(tinh_chu, tach_phanchinh(tinhmoi)),
                fuzz.token_sort_ratio(tinh_chu, tach_phanchinh(tinhmoi)) * 0.9
            )
            
            if score_tinh_moi < FUZZY_THRESHOLDS['tinh_moi_min']:
                continue

            total_score = 0.7 * score_xa_moi + 0.3 * score_tinh_moi
            
            if total_score > best_score:
                best_score = total_score
                best_match_index = i
                best_sheet = 'sheet1'
        
        # Check sheet2
        for i, item in enumerate(self.mapping_sheet2):
            apcu, xacu, huyencu, tinhcu, apmoi, xamoi, tinhmoi = item
            xa_moi_chu, xa_moi_so = tach_chu_so(tach_phanchinh(xamoi))
            
            # Lọc sơ bộ theo số
            if xa_so and xa_so != xa_moi_so:
                continue
            
            # Tính điểm cho xã mới
            score_xa_moi = max(
                fuzz.ratio(xa_chu, xa_moi_chu),
                fuzz.partial_ratio(xa_chu, xa_moi_chu),
                fuzz.token_sort_ratio(xa_chu, xa_moi_chu) * 0.9
            )
            
            if score_xa_moi < FUZZY_THRESHOLDS['xa_moi_min']:
                continue

            # Tính điểm cho tỉnh mới
            score_tinh_moi = max(
                fuzz.ratio(tinh_chu, tach_phanchinh(tinhmoi)),
                fuzz.partial_ratio(tinh_chu, tach_phanchinh(tinhmoi)),
                fuzz.token_sort_ratio(tinh_chu, tach_phanchinh(tinhmoi)) * 0.9
            )
            
            if score_tinh_moi < FUZZY_THRESHOLDS['tinh_moi_min']:
                continue

            total_score = 0.7 * score_xa_moi + 0.3 * score_tinh_moi
            
            if total_score > best_score:
                best_score = total_score
                best_match_index = i
                best_sheet = 'sheet2'

        if best_match_index is not None and best_score >= FUZZY_THRESHOLDS['total_min']:
            if best_sheet == 'sheet1':
                # FIXED: Return original data
                original_item = self.mapping_sheet1_original[best_match_index]
                return original_item + ('Địa chỉ mới - giữ nguyên',)
            else:  # sheet2
                # FIXED: Return original data, convert to sheet1 format
                original_item = self.mapping_sheet2_original[best_match_index]
                return (original_item[1], original_item[2], original_item[3], 
                       original_item[5], original_item[6], 'Địa chỉ mới - giữ nguyên')
        
        return (None, None, None, None, None, 'Không tìm thấy trong danh sách địa chỉ (thiếu huyện)')
    
    def _match_full_address(self, xa_chu, xa_so, huyen_chu, tinh_chu, xa_orig, huyen_orig, tinh_orig):
        """Match với địa chỉ đầy đủ - FIXED to return original data"""
        # Tìm kiếm chính xác trước
        exact_key = (xa_chu + ' ' + xa_so if xa_so else xa_chu, huyen_chu, tinh_chu)
        if exact_key in self.cache:
            matches = self.cache[exact_key]
            # Ưu tiên sheet1 nếu có
            for index, sheet in matches:
                if sheet == 'sheet1':
                    # FIXED: Return original data
                    original_item = self.mapping_sheet1_original[index]
                    return original_item + ('',)
            # Nếu không có trong sheet1, lấy từ sheet2
            if matches:
                index, sheet = matches[0]
                if sheet == 'sheet2':
                    # FIXED: Return original data, convert to sheet1 format
                    original_item = self.mapping_sheet2_original[index]
                    return (original_item[1], original_item[2], original_item[3], 
                           original_item[5], original_item[6], 'Xã cấu véo')

        # Nếu không tìm thấy chính xác, dùng fuzzy matching
        return self._fuzzy_match_full_address(xa_chu, xa_so, huyen_chu, tinh_chu, xa_orig, huyen_orig, tinh_orig)
    
    def _fuzzy_match_full_address(self, xa_chu, xa_so, huyen_chu, tinh_chu, xa_orig, huyen_orig, tinh_orig):
        """Fuzzy match với địa chỉ đầy đủ - FIXED to return original data"""
        best_score = 0
        best_match_index = None
        best_sheet = None
        
        # Tạo danh sách ứng viên
        candidates = []
        
        # Check sheet1
        for i, item in enumerate(self.mapping_sheet1):
            xacu, huyencu, tinhcu, xamoi, tinhmoi = item
            xacu_chu, xacu_so = tach_chu_so(tach_phanchinh(xacu))
            
            # Lọc sơ bộ theo số
            if xa_so and xa_so != xacu_so:
                continue
            
            # Lọc sơ bộ theo độ dài tên xã
            if abs(len(xa_chu) - len(xacu_chu)) > 5:
                continue
                
            candidates.append((i, 'sheet1'))
        
        # Check sheet2
        for i, item in enumerate(self.mapping_sheet2):
            apcu, xacu, huyencu, tinhcu, apmoi, xamoi, tinhmoi = item
            xacu_chu, xacu_so = tach_chu_so(tach_phanchinh(xacu))
            
            # Lọc sơ bộ theo số
            if xa_so and xa_so != xacu_so:
                continue
            
            # Lọc sơ bộ theo độ dài tên xã
            if abs(len(xa_chu) - len(xacu_chu)) > 5:
                continue
                
            candidates.append((i, 'sheet2'))

        # Tính điểm cho các ứng viên
        for index, sheet in candidates:
            if sheet == 'sheet1':
                item = self.mapping_sheet1[index]
                xacu, huyencu, tinhcu, xamoi, tinhmoi = item
            else:  # sheet2
                item = self.mapping_sheet2[index]
                apcu, xacu, huyencu, tinhcu, apmoi, xamoi, tinhmoi = item
            
            xacu_chu, xacu_so = tach_chu_so(tach_phanchinh(xacu))
            
            # Tính điểm xã với các phương pháp khác nhau
            score_xa_ratio = fuzz.ratio(xa_chu, xacu_chu)
            score_xa_partial = fuzz.partial_ratio(xa_chu, xacu_chu)
            score_xa_token = fuzz.token_sort_ratio(xa_chu, xacu_chu)
            score_xa = max(score_xa_ratio, score_xa_partial, score_xa_token * 0.9)
            
            if score_xa < FUZZY_THRESHOLDS['xa_min']:
                continue

            # Tính điểm huyện
            score_huyen = max(
                fuzz.ratio(huyen_chu, tach_phanchinh(huyencu)),
                fuzz.partial_ratio(huyen_chu, tach_phanchinh(huyencu)),
                fuzz.token_sort_ratio(huyen_chu, tach_phanchinh(huyencu)) * 0.9
            )
            if score_huyen < FUZZY_THRESHOLDS['huyen_min']:
                continue

            # Tính điểm tỉnh
            score_tinh = max(
                fuzz.ratio(tinh_chu, tach_phanchinh(tinhcu)),
                fuzz.partial_ratio(tinh_chu, tach_phanchinh(tinhcu)),
                fuzz.token_sort_ratio(tinh_chu, tach_phanchinh(tinhcu)) * 0.9
            )
            if score_tinh < FUZZY_THRESHOLDS['tinh_min']:
                continue

            # Điều chỉnh trọng số
            total_score = (
                0.5 * score_xa +      # Xã quan trọng nhất
                0.3 * score_huyen +   # Huyện quan trọng thứ hai
                0.2 * score_tinh      # Tỉnh ít quan trọng nhất
            )
            
            if total_score > best_score:
                best_score = total_score
                best_match_index = index
                best_sheet = sheet

        # Ngưỡng điểm tối thiểu
        if best_match_index is not None and best_score >= FUZZY_THRESHOLDS['total_min']:
            if best_sheet == 'sheet1':
                # FIXED: Return original data
                original_item = self.mapping_sheet1_original[best_match_index]
                return original_item + ('',)
            else:  # sheet2
                # FIXED: Return original data, convert to sheet1 format
                original_item = self.mapping_sheet2_original[best_match_index]
                return (original_item[1], original_item[2], original_item[3], 
                       original_item[5], original_item[6], 'Xã cấu véo')

        # Nếu không match được, kiểm tra các trường hợp lỗi
        return self._check_error_cases(xa_orig, huyen_orig, tinh_orig)
    
    def _check_error_cases(self, xa, huyen, tinh):
        """Kiểm tra các trường hợp lỗi cụ thể - using normalized data for checking"""
        # Use normalized data for error checking
        all_mapping = self.mapping_sheet1 + [
            (item[1], item[2], item[3], item[5], item[6]) for item in self.mapping_sheet2
        ]
        
        xa_found = any(
            tach_chu_so(tach_phanchinh(xa))[0] == tach_chu_so(tach_phanchinh(m[0]))[0]
            for m in all_mapping
        )
        huyen_found = any(
            tach_phanchinh(huyen) == tach_phanchinh(m[1])
            for m in all_mapping
        )
        tinh_found = any(
            tach_phanchinh(tinh) == tach_phanchinh(m[2])
            for m in all_mapping
        )

        if not xa_found:
            return (None, None, None, None, None, 'Thông tin sai: Xã không tồn tại')

        if not huyen_found:
            return (None, None, None, None, None, 'Thông tin sai: Huyện không tồn tại')

        if not tinh_found:
            return (None, None, None, None, None, 'Thông tin sai: Tỉnh không tồn tại')

        return (None, None, None, None, None, 'Thông tin không khớp')


# Global matcher instance
fuzzy_matcher = FuzzyMatcher()


def fuzzy_match_row(xa, huyen, tinh, ap=None, address_detail=None):
    """
    Wrapper function để maintain compatibility với code cũ - UPDATED
    """
    return fuzzy_matcher.match_row(xa, huyen, tinh, ap, address_detail)