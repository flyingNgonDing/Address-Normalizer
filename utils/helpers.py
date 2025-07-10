"""
Helper functions và utilities
WINDOWS VERSION - Optimized for Windows 8/10/11
FIXED: get_mapping_file_path() now ALWAYS returns external path
"""
import os
import sys
import subprocess
from pathlib import Path


def format_time(seconds):
    """
    Format thời gian từ seconds thành MM:SS
    
    Args:
        seconds: Số giây
        
    Returns:
        str: Thời gian định dạng MM:SS
    """
    if seconds < 0:
        seconds = 0
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    return f"{minutes:02d}:{secs:02d}"


def format_number(number):
    """
    Format số với dấu phẩy phân cách hàng nghìn
    
    Args:
        number: Số cần format
        
    Returns:
        str: Số đã format
    """
    return f"{number:,}"


def open_file_with_system(file_path):
    """
    Mở file bằng ứng dụng mặc định của Windows
    
    Args:
        file_path: Đường dẫn file
        
    Returns:
        bool: True nếu thành công
    """
    if not os.path.exists(file_path):
        return False
    
    try:
        # Windows specific methods
        if sys.platform.startswith('win'):
            # Try multiple methods for better compatibility
            try:
                # Method 1: os.startfile (preferred for Windows)
                os.startfile(file_path)
                return True
            except Exception:
                try:
                    # Method 2: subprocess with start command
                    subprocess.run(['start', '', file_path], shell=True, check=True)
                    return True
                except Exception:
                    try:
                        # Method 3: explorer.exe
                        subprocess.run(['explorer.exe', file_path], check=True)
                        return True
                    except Exception:
                        return False
        else:
            # Fallback for non-Windows (shouldn't happen in this version)
            if sys.platform.startswith('darwin'):  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux
                subprocess.call(['xdg-open', file_path])
            return True
    except Exception as e:
        print(f"Error opening file: {e}")
        return False


def get_mapping_file_path():
    """
    Lấy đường dẫn file mapping.xlsx - FIXED for EXTERNAL mapping
    
    CRITICAL: This function MUST return the external mapping.xlsx path,
    NOT the internal/bundled path, to ensure users can update mapping files.
    
    Returns:
        str: Đường dẫn file mapping external
    """
    if getattr(sys, 'frozen', False):
        # ✅ RUNNING AS EXECUTABLE: Find mapping.xlsx next to .exe file
        # sys.executable gives path to the .exe file
        exe_dir = Path(sys.executable).parent
        mapping_path = exe_dir / "mapping.xlsx"
        
        print(f"🔍 [EXECUTABLE MODE] Looking for external mapping.xlsx at: {mapping_path}")
        print(f"🔍 [EXECUTABLE MODE] File exists: {mapping_path.exists()}")
        return str(mapping_path)
    else:
        # ✅ RUNNING IN DEVELOPMENT: Find mapping.xlsx in project root
        # __file__ gives path to this helpers.py file
        project_root = Path(__file__).parent.parent  # Go up from utils/ to project root
        mapping_path = project_root / "mapping.xlsx"
        
        print(f"🔍 [DEVELOPMENT MODE] Looking for mapping.xlsx at: {mapping_path}")
        print(f"🔍 [DEVELOPMENT MODE] File exists: {mapping_path.exists()}")
        return str(mapping_path)


def get_mapping_file_path_old():
    """
    OLD VERSION - kept for reference
    Lấy đường dẫn file mapping.xlsx - WINDOWS VERSION
    
    Returns:
        str: Đường dẫn file mapping
    """
    from config import APPLICATION_PATH
    
    # Use pathlib for better Windows path handling
    app_path = Path(APPLICATION_PATH)
    
    # Thử nhiều vị trí có thể có
    possible_paths = [
        app_path / "mapping.xlsx",
        app_path / "data" / "mapping.xlsx",
        app_path / "resources" / "mapping.xlsx",
        app_path.parent / "mapping.xlsx",
        app_path.parent / "data" / "mapping.xlsx",
    ]
    
    for path in possible_paths:
        path_str = str(path.resolve())
        print(f"[DEBUG] Checking mapping file at: {path_str}")
        if path.exists():
            print(f"[DEBUG] Found mapping file at: {path_str}")
            return path_str
    
    # Fallback: return first path anyway (in same directory as exe)
    fallback_path = str((app_path / "mapping.xlsx").resolve())
    print(f"[DEBUG] Mapping file not found, using fallback: {fallback_path}")
    return fallback_path


def ensure_directory_exists(file_path):
    """
    Đảm bảo thư mục chứa file tồn tại
    
    Args:
        file_path: Đường dẫn file
    """
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)


def is_file_locked(file_path):
    """
    Kiểm tra file có đang bị lock không (Windows version)
    
    Args:
        file_path: Đường dẫn file
        
    Returns:
        bool: True nếu file bị lock
    """
    if not os.path.exists(file_path):
        return False
    
    try:
        # Windows specific file lock check
        if sys.platform.startswith('win'):
            # Try to open file in exclusive mode
            import msvcrt
            with open(file_path, 'r+b') as f:
                try:
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                    return False
                except IOError:
                    return True
        else:
            # Fallback method
            with open(file_path, 'a'):
                pass
            return False
    except (IOError, OSError):
        return True


def sanitize_filename(filename):
    """
    Làm sạch tên file, loại bỏ ký tự không hợp lệ cho Windows
    
    Args:
        filename: Tên file gốc
        
    Returns:
        str: Tên file đã làm sạch
    """
    import re
    
    # Windows specific invalid characters
    invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
    clean_name = re.sub(invalid_chars, '_', filename)
    
    # Remove leading/trailing dots and spaces (Windows restriction)
    clean_name = clean_name.strip('. ')
    
    # Windows reserved names
    reserved_names = [
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ]
    
    name_without_ext = os.path.splitext(clean_name)[0].upper()
    if name_without_ext in reserved_names:
        clean_name = f"_{clean_name}"
    
    # Loại bỏ dấu cách thừa
    clean_name = ' '.join(clean_name.split())
    
    # Giới hạn độ dài tên file (Windows limit: 255 characters)
    if len(clean_name) > 200:  # Leave some room for path
        name, ext = os.path.splitext(clean_name)
        clean_name = name[:200-len(ext)] + ext
    
    return clean_name


def get_windows_version():
    """
    Lấy thông tin phiên bản Windows
    
    Returns:
        dict: Thông tin version Windows
    """
    try:
        import platform
        version_info = {
            'system': platform.system(),
            'release': platform.release(),
            'version': platform.version(),
            'machine': platform.machine(),
            'processor': platform.processor(),
        }
        
        # Detect specific Windows versions
        if sys.platform.startswith('win'):
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                   r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                try:
                    version_info['build'] = winreg.QueryValueEx(key, "CurrentBuild")[0]
                    version_info['product_name'] = winreg.QueryValueEx(key, "ProductName")[0]
                except:
                    pass
                winreg.CloseKey(key)
            except:
                pass
        
        return version_info
    except Exception as e:
        print(f"Error getting Windows version: {e}")
        return {'system': 'Unknown', 'error': str(e)}


def check_windows_compatibility():
    """
    Kiểm tra tính tương thích với Windows
    
    Returns:
        dict: Kết quả kiểm tra tương thích
    """
    result = {
        'compatible': True,
        'warnings': [],
        'version_info': get_windows_version()
    }
    
    if not sys.platform.startswith('win'):
        result['compatible'] = False
        result['warnings'].append("Không phải hệ điều hành Windows")
        return result
    
    # Check Windows version
    try:
        version = result['version_info'].get('release', '')
        if version in ['8', '8.1', '10', '11']:
            result['warnings'].append(f"Windows {version} - Tương thích tốt")
        elif version == '7':
            result['warnings'].append("Windows 7 - Có thể gặp một số vấn đề tương thích")
        else:
            result['warnings'].append(f"Windows {version} - Chưa được kiểm tra đầy đủ")
    except:
        result['warnings'].append("Không thể xác định phiên bản Windows")
    
    return result


def setup_windows_environment():
    """
    Thiết lập môi trường Windows tối ưu
    """
    try:
        # Set DPI awareness for Windows 8.1+
        if sys.platform.startswith('win'):
            try:
                import ctypes
                # Try to set DPI awareness
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except:
                try:
                    # Fallback for older Windows versions
                    ctypes.windll.user32.SetProcessDPIAware()
                except:
                    pass
    except Exception as e:
        print(f"Warning: Could not set up Windows environment: {e}")


def get_resource_path(relative_path):
    """
    Get absolute path to resource - UPDATED to prevent mapping.xlsx bundling
    
    This function is for OTHER resources that should be bundled,
    but mapping.xlsx should NEVER use this function.
    
    Args:
        relative_path: Relative path to resource
        
    Returns:
        str: Absolute path to resource
    """
    # ⚠️ WARNING: Do not use this function for mapping.xlsx!
    if 'mapping.xlsx' in str(relative_path).lower():
        raise ValueError("❌ mapping.xlsx must use get_mapping_file_path(), not get_resource_path()!")
    
    if getattr(sys, 'frozen', False):
        # Running as executable - look in temp directory for bundled resources
        base_path = sys._MEIPASS
    else:
        # Running in development
        base_path = os.path.dirname(__file__)
    
    return os.path.join(base_path, relative_path)