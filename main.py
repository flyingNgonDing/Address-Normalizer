#!/usr/bin/env python3
"""
Chương trình chuẩn hóa địa chỉ bệnh nhân
WINDOWS VERSION - ENHANCED with File Watcher for Auto-restart + EMBEDDED ICON + XLS SUPPORT
Entry point cho ứng dụng
"""
import sys
import os
import traceback
from pathlib import Path

# ===== PYINSTALLER PATH FIX =====
def setup_paths():
    """Setup paths for PyInstaller"""
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.executable)))
        application_path = bundle_dir
    else:
        # Running in development
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    # Add to Python path
    if application_path not in sys.path:
        sys.path.insert(0, application_path)
    
    # Add subdirectories to path - CORRECTED based on your structure
    subdirs = ['core', 'data', 'gui', 'utils']
    for subdir in subdirs:
        subdir_path = os.path.join(application_path, subdir)
        if os.path.exists(subdir_path) and subdir_path not in sys.path:
            sys.path.insert(0, subdir_path)
    
    print(f"Application path: {application_path}")
    print(f"Python paths: {sys.path[:5]}...")  # Show first 5 paths
    
    return application_path

# Setup paths first
APPLICATION_PATH = setup_paths()

# ===== SIMPLE ICON MANAGEMENT =====
def setup_window_icon(root):
    """Setup window icon - PyInstaller embedded version"""
    try:
        if getattr(sys, 'frozen', False):
            # Running as compiled exe - icon is embedded by PyInstaller
            # Try to find icon in PyInstaller temp directory
            bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.executable)))
            icon_path = Path(bundle_dir) / "icon.ico"
            
            if icon_path.exists():
                root.iconbitmap(str(icon_path))
                print(f"✅ Embedded icon loaded: {icon_path}")
                return True
            else:
                print(f"⚠️ Embedded icon not found in: {bundle_dir}")
                return False
        else:
            # Running in development - look for icon.ico in project folder
            dev_locations = [
                Path(__file__).parent / "icon.ico",
                Path(APPLICATION_PATH) / "icon.ico",
                Path.cwd() / "icon.ico"
            ]
            
            for icon_path in dev_locations:
                if icon_path.exists():
                    root.iconbitmap(str(icon_path))
                    print(f"✅ Development icon loaded: {icon_path}")
                    return True
            
            print(f"⚠️ Development icon not found")
            return False
            
    except Exception as e:
        print(f"⚠️ Could not set window icon: {e}")
        return False

# ===== IMPORTS WITH ERROR HANDLING =====
def safe_import():
    """Import modules safely with detailed error reporting"""
    try:
        # Test basic imports first
        print("Testing basic imports...")
        import tkinter as tk
        print("✅ tkinter OK")
        
        # Test pandas
        import pandas as pd
        print("✅ pandas OK")
        
        # Test GUI modules - CORRECTED import paths
        print("Testing GUI imports...")
        try:
            from gui.main_window import MainWindow
            print("✅ gui.main_window OK")
        except ImportError as e:
            print(f"❌ gui.main_window failed: {e}")
            # Try alternative import
            from main_window import MainWindow
            print("✅ main_window (direct) OK")
        
        try:
            from gui.styles import setup_ui_styles, apply_windows_theme
            print("✅ gui.styles OK")
        except ImportError as e:
            print(f"❌ gui.styles failed: {e}")
            from styles import setup_ui_styles, apply_windows_theme
            print("✅ styles (direct) OK")
        
        try:
            from gui.drag_drop import DragDropHandler, is_drag_drop_available
            print("✅ gui.drag_drop OK")
        except ImportError as e:
            print(f"❌ gui.drag_drop failed: {e}")
            from drag_drop import DragDropHandler, is_drag_drop_available
            print("✅ drag_drop (direct) OK")
        
        # Test other modules - CORRECTED import paths
        print("Testing other imports...")
        try:
            from data.mapping_loader import load_mapping
            print("✅ data.mapping_loader OK")
        except ImportError as e:
            print(f"❌ data.mapping_loader failed: {e}")
            from mapping_loader import load_mapping
            print("✅ mapping_loader (direct) OK")
        
        try:
            from utils.helpers import check_windows_compatibility
            print("✅ utils.helpers OK")
        except ImportError as e:
            print(f"❌ utils.helpers failed: {e}")
            from helpers import check_windows_compatibility
            print("✅ helpers (direct) OK")
        
        # NEW: Import file watcher
        print("Testing file watcher...")
        try:
            from gui.file_watcher import start_file_watching, stop_file_watching, is_file_watching_active
            print("✅ gui.file_watcher OK")
        except ImportError as e:
            print(f"❌ gui.file_watcher failed: {e}")
            try:
                from file_watcher import start_file_watching, stop_file_watching, is_file_watching_active
                print("✅ file_watcher (direct) OK")
            except ImportError as e2:
                print(f"❌ file_watcher (direct) failed: {e2}")
                # Fallback functions if file_watcher not available
                def start_file_watching(main_window):
                    print("⚠️ File watcher not available")
                    return False
                def stop_file_watching():
                    pass
                def is_file_watching_active():
                    return False
                print("✅ file_watcher (fallback) OK")
        
        # Import config
        import config
        print("✅ config OK")
        
        return {
            'MainWindow': MainWindow,
            'setup_ui_styles': setup_ui_styles,
            'apply_windows_theme': apply_windows_theme,
            'load_mapping': load_mapping,
            'check_windows_compatibility': check_windows_compatibility,
            'start_file_watching': start_file_watching,
            'stop_file_watching': stop_file_watching,
            'is_file_watching_active': is_file_watching_active,
            'config': config
        }
        
    except Exception as e:
        print(f"❌ Import error: {e}")
        print(f"❌ Traceback: {traceback.format_exc()}")
        return None

def check_python_version():
    """Kiểm tra phiên bản Python"""
    if sys.version_info < (3, 7):
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror(
                "Lỗi phiên bản Python", 
                f"Chương trình yêu cầu Python 3.7 trở lên.\n"
                f"Phiên bản hiện tại: {sys.version}\n\n"
                f"Vui lòng cập nhật Python."
            )
        except:
            print(f"❌ Python version {sys.version_info} not supported. Need 3.7+")
        return False
    return True

def check_dependencies():
    """Kiểm tra các dependencies cần thiết - ENHANCED for .xls support"""
    required_modules = [
        ('pandas', 'pandas'),
        ('thefuzz', 'thefuzz'),
        ('openpyxl', 'openpyxl'),
        ('xlrd', 'xlrd'),  # For .xls support
    ]
    
    # Optional modules for enhanced features
    optional_modules = [
        ('watchdog', 'watchdog'),  # For file watching
        ('tkinterdnd2', 'tkinterdnd2'),  # For drag & drop
    ]
    
    missing_modules = []
    missing_optional = []
    version_warnings = []
    
    for module_name, import_name in required_modules:
        try:
            module = __import__(import_name)
            version = getattr(module, '__version__', 'unknown')
            
            # Special version checks
            if module_name == 'xlrd':
                # Check xlrd version for .xls support
                if version != 'unknown':
                    try:
                        version_parts = version.split('.')
                        major = int(version_parts[0])
                        minor = int(version_parts[1]) if len(version_parts) > 1 else 0
                        
                        if major > 2 or (major == 2 and minor > 0):
                            version_warnings.append(f"xlrd {version} may not support .xls files. Recommended: xlrd==2.0.1")
                            print(f"⚠️ {module_name} {version} (may not support .xls)")
                        else:
                            print(f"✅ {module_name} {version} (good for .xls support)")
                    except:
                        version_warnings.append(f"xlrd {version} - version unclear, recommend xlrd==2.0.1")
                        print(f"⚠️ {module_name} {version} (version unclear)")
                else:
                    version_warnings.append(f"xlrd version unknown - recommend xlrd==2.0.1")
                    print(f"⚠️ {module_name} (version unknown)")
            elif module_name == 'pandas':
                print(f"✅ {module_name} {version}")
                # Check pandas version
                if version != 'unknown':
                    try:
                        version_parts = version.split('.')
                        major = int(version_parts[0])
                        minor = int(version_parts[1]) if len(version_parts) > 1 else 0
                        
                        if major < 1 or (major == 1 and minor < 3):
                            version_warnings.append(f"pandas {version} is old. Recommended: pandas>=1.3.0")
                    except:
                        pass
            else:
                print(f"✅ {module_name} {version}")
                
        except ImportError:
            missing_modules.append(module_name)
            print(f"❌ {module_name}")
    
    for module_name, import_name in optional_modules:
        try:
            module = __import__(import_name)
            version = getattr(module, '__version__', 'unknown')
            print(f"✅ {module_name} {version} (optional)")
        except ImportError:
            missing_optional.append(module_name)
            
            if module_name == 'watchdog':
                print(f"⚠️ {module_name} (optional) - Auto-restart feature disabled")
            elif module_name == 'tkinterdnd2':
                print(f"⚠️ {module_name} (optional) - Drag & drop feature disabled")
    
    # Show warnings
    if version_warnings:
        print("\n⚠️ VERSION WARNINGS:")
        for warning in version_warnings:
            print(f"   {warning}")
    
    if missing_modules:
        error_msg = (
            f"❌ Thiếu các thư viện bắt buộc:\n"
            f"{', '.join(missing_modules)}\n\n"
        )
        
        if 'xlrd' in missing_modules:
            error_msg += (
                f"🔧 ĐẶC BIỆT CHO .XLS SUPPORT:\n"
                f"pip install xlrd==2.0.1\n\n"
            )
        
        error_msg += (
            f"Ứng dụng có thể không hoạt động đúng.\n\n"
            f"Vui lòng chạy:\n"
            f"pip install {' '.join(missing_modules)}\n\n"
        )
        
        if 'xlrd' in missing_modules:
            error_msg += (
                f"📝 LỆNH CỤ THỂ:\n"
                f"pip install xlrd==2.0.1 openpyxl pandas thefuzz\n\n"
            )
        
        print(error_msg)
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Thiếu thư viện", error_msg)
        except:
            pass
        return False
    
    if version_warnings:
        warning_msg = (
            f"⚠️ Cảnh báo về phiên bản:\n\n"
            f"{chr(10).join(version_warnings)}\n\n"
            f"🔧 LỆNH KHẮC PHỤC:\n"
            f"pip uninstall xlrd -y\n"
            f"pip install xlrd==2.0.1\n\n"
            f"Hoặc cài đặt lại tất cả:\n"
            f"pip install xlrd==2.0.1 openpyxl pandas>=1.3.0 thefuzz\n\n"
            f"Bạn có muốn tiếp tục không?"
        )
        
        print(warning_msg)
        try:
            import tkinter.messagebox as messagebox
            result = messagebox.askyesno("Cảnh báo phiên bản", warning_msg)
            if not result:
                return False
        except:
            pass
    
    if missing_optional:
        print(f"\n💡 Tính năng bổ sung (tùy chọn):")
        for module in missing_optional:
            if module == 'watchdog':
                print(f"   pip install watchdog  # Auto-restart khi mapping.xlsx thay đổi")
            elif module == 'tkinterdnd2':
                print(f"   pip install tkinterdnd2  # Drag & drop files")
    
    return True

def test_excel_file_support():
    """Test Excel file support explicitly"""
    print("\n🔍 Testing Excel file support...")
    
    engines = {}
    
    # Test openpyxl
    try:
        import openpyxl
        engines['openpyxl'] = f"✅ Available (version: {openpyxl.__version__})"
        print(f"✅ openpyxl: {openpyxl.__version__} (.xlsx support)")
    except ImportError:
        engines['openpyxl'] = "❌ Not available"
        print("❌ openpyxl: Not available (.xlsx support DISABLED)")
    
    # Test xlrd
    try:
        import xlrd
        engines['xlrd'] = f"✅ Available (version: {xlrd.__version__})"
        print(f"✅ xlrd: {xlrd.__version__} (.xls support)")
        
        # Test if xlrd can actually read .xls files
        version = xlrd.__version__
        version_parts = version.split('.')
        major = int(version_parts[0])
        minor = int(version_parts[1]) if len(version_parts) > 1 else 0
        
        if major > 2 or (major == 2 and minor > 0):
            print(f"⚠️ xlrd {version} may not support .xls files!")
            print(f"💡 Recommended: pip install xlrd==2.0.1")
        else:
            print(f"✅ xlrd {version} should support .xls files")
            
    except ImportError:
        engines['xlrd'] = "❌ Not available"
        print("❌ xlrd: Not available (.xls support DISABLED)")
        print("💡 Install with: pip install xlrd==2.0.1")
    
    # Test pandas Excel integration
    try:
        import pandas as pd
        print(f"✅ pandas: {pd.__version__} (Excel integration)")
        
        # Test if pandas can use the engines
        available_engines = []
        if 'openpyxl' in engines and '✅' in engines['openpyxl']:
            available_engines.append('openpyxl')
        if 'xlrd' in engines and '✅' in engines['xlrd']:
            available_engines.append('xlrd')
        
        if available_engines:
            print(f"✅ Available engines: {', '.join(available_engines)}")
        else:
            print("❌ No Excel engines available!")
            
    except ImportError:
        print("❌ pandas: Not available")
    
    return engines

def setup_windows_environment():
    """Thiết lập môi trường Windows"""
    try:
        # Basic Windows setup
        if sys.platform.startswith('win'):
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except:
                pass
                
    except Exception as e:
        print(f"Warning: Could not setup Windows environment: {e}")

def create_tkinter_root():
    """Tạo Tkinter root window với Windows optimization và EMBEDDED ICON + DRAG & DROP"""
    try:
        import tkinter as tk
        
        # ✅ FIXED: Always try to use TkinterDnD for drag & drop support
        try:
            from tkinterdnd2 import TkinterDnD
            root = TkinterDnD.Tk()
            print("✅ Using TkinterDnD for drag & drop support (including frozen executable)")
        except ImportError:
            print("⚠️ TkinterDnD not available, using standard Tkinter")
            root = tk.Tk()
            print("✅ Using standard Tkinter (no drag & drop)")
        
        # Test basic tkinter functionality
        root.withdraw()  # Hide window temporarily
        
        # ✅ Setup window icon IMMEDIATELY after creating root
        print("🎨 Setting up window icon...")
        icon_success = setup_window_icon(root)
        if icon_success:
            print("✅ Window icon set successfully")
        else:
            print("⚠️ Using default system icon")
        
        # Windows specific root configuration
        if sys.platform.startswith('win'):
            try:
                # Set Windows-specific attributes
                root.wm_attributes('-alpha', 1.0)
                
                # Try to set DPI awareness
                try:
                    import ctypes
                    ctypes.windll.shcore.SetProcessDpiAwareness(1)
                except:
                    pass
                    
            except Exception as e:
                print(f"Warning: Could not set Windows attributes: {e}")
        
        return root
        
    except Exception as e:
        error_msg = f"❌ Không thể khởi tạo giao diện:\n{str(e)}"
        print(error_msg)
        raise

def main():
    """Main function - ENHANCED with Excel support testing + XLS support"""
    print("🚀 Khởi động ứng dụng Chuẩn hóa địa chỉ bệnh nhân")
    print("=" * 50)
    
    # Global variable to store file watcher functions
    file_watcher_funcs = None
    
    try:
        # 1. Check Python version
        print("🐍 Checking Python version...")
        if not check_python_version():
            input("Press Enter to exit...")
            sys.exit(1)
        
        # 2. Setup paths and imports
        print("📥 Setting up imports...")
        modules = safe_import()
        if not modules:
            print("❌ Failed to import required modules")
            input("Press Enter to exit...")
            sys.exit(1)
        
        # Extract modules
        MainWindow = modules['MainWindow']
        setup_ui_styles = modules['setup_ui_styles'] 
        apply_windows_theme = modules['apply_windows_theme']
        load_mapping = modules['load_mapping']
        check_windows_compatibility = modules['check_windows_compatibility']
        
        # NEW: Extract file watcher functions
        start_file_watching = modules['start_file_watching']
        stop_file_watching = modules['stop_file_watching']
        is_file_watching_active = modules['is_file_watching_active']
        file_watcher_funcs = (start_file_watching, stop_file_watching, is_file_watching_active)
        
        # 3. Check dependencies
        print("📦 Checking dependencies...")
        if not check_dependencies():
            input("Press Enter to exit...")
            sys.exit(1)
        
        # 4. Test Excel file support specifically
        test_excel_file_support()
        
        # 5. Setup Windows environment
        print("🪟 Setting up Windows environment...")
        setup_windows_environment()
        
        # 6. Create Tkinter root - WITH EMBEDDED ICON
        print("🖼️  Creating GUI window with embedded icon...")
        root = create_tkinter_root()
        
        # 7. Load mapping data
        print("📊 Loading mapping data...")
        try:
            load_mapping()
            print("✅ Mapping data loaded successfully.")
        except Exception as e:
            error_msg = f"❌ Lỗi load dữ liệu mapping:\n{str(e)}\n\nKiểm tra file mapping.xlsx có tồn tại không."
            print(error_msg)
            try:
                import tkinter.messagebox as messagebox
                messagebox.showerror("Lỗi dữ liệu", error_msg)
            except:
                pass
            # Continue without mapping data
            print("⚠️  Continuing without mapping data...")
        
        # 8. Show window and create main window
        print("🎨 Creating main window...")
        root.deiconify()  # Show window
        app = MainWindow(root)
        
        # ✅ Re-apply icon after MainWindow is created (in case it gets reset)
        print("🔄 Re-applying window icon...")
        setup_window_icon(root)
        
        # 9. NEW: Start file watching for auto-restart
        print("👁️ Setting up file watching...")
        file_watch_success = start_file_watching(app)
        if file_watch_success:
            print("✅ Auto-restart enabled - mapping.xlsx changes will trigger restart")
        else:
            print("⚠️ Auto-restart disabled - manual restart required for mapping changes")
        
        # 10. Setup cleanup on exit
        def on_app_exit():
            """Cleanup function when app exits"""
            print("🧹 Cleaning up...")
            if file_watcher_funcs:
                try:
                    stop_file_watching()
                except Exception as e:
                    print(f"Warning: Error stopping file watcher: {e}")
            root.quit()
            root.destroy()
        
        # Bind cleanup to window close
        root.protocol("WM_DELETE_WINDOW", on_app_exit)
        
        # 11. Start main loop
        print("✅ Application started successfully!")
        print("🖱️  GUI should be visible now...")
        print("")
        print("🆕 NEW FEATURES:")
        print("   ✅ Settings button is now visible")
        print("   ✅ Support for .xls files (requires xlrd==2.0.1)")
        print("   ✅ Multi-sheet Excel processing")
        print("   ✅ Enhanced ấp matching for sheet2")
        print("   ✅ Author info displayed in bottom-right")
        print("   ✅ Embedded icon (no external files needed)")
        if file_watch_success:
            print("   ✅ Auto-restart when mapping.xlsx is saved")
        else:
            print("   ⚠️ Auto-restart disabled (watchdog not available)")
        print("")
        print("🔧 XLS SUPPORT REQUIREMENTS:")
        print("   CRITICAL: pip install xlrd==2.0.1")
        print("   Note: xlrd versions > 2.0.1 do NOT support .xls files")
        print("   If you get errors with .xls files, run:")
        print("     pip uninstall xlrd -y")
        print("     pip install xlrd==2.0.1")
        print("")
        print("📝 MAPPING.XLSX SHEET2 NEW STRUCTURE:")
        print("   Cột A: apcu (ấp cũ)")
        print("   Cột B: xacu (xã cũ)")
        print("   Cột C: huyencu (huyện cũ)")
        print("   Cột D: tinhcu (tỉnh cũ)")
        print("   Cột E: apmoi (ấp mới)")
        print("   Cột F: xamoi (xã mới)")
        print("   Cột G: tinhmoi (tỉnh mới)")
        print("")
        print("🎯 AUTHOR INFO:")
        print("   Viện Pasteur TP HCM - Khoa KSPNBT - Nhóm TKBC")
        print("   Tác giả: Đinh Văn Ngôn (08888 31135)")
        print("")
        
        root.mainloop()
        
    except KeyboardInterrupt:
        print("❌ Application interrupted by user")
        if file_watcher_funcs:
            try:
                file_watcher_funcs[1]()  # stop_file_watching
            except:
                pass
        sys.exit(0)
        
    except Exception as e:
        error_msg = f"❌ Lỗi khởi tạo ứng dụng:\n{str(e)}\n\nStacktrace:\n{traceback.format_exc()}"
        print(error_msg)
        
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Lỗi", error_msg)
        except:
            print("❌ Không thể hiển thị messagebox")
        
        # Cleanup on error
        if file_watcher_funcs:
            try:
                file_watcher_funcs[1]()  # stop_file_watching
            except:
                pass
        
        input("Press Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    # Windows specific startup checks
    if sys.platform.startswith('win'):
        # Set console title for debugging
        try:
            os.system("title PIHCM - Chuẩn hóa địa chỉ bệnh nhân - XLS SUPPORT")
        except:
            pass
        
        # Print system info
        print(f"🖥️  Running on: {sys.platform}")
        print(f"📁 Working directory: {os.getcwd()}")
        print(f"🐍 Python: {sys.version}")
        print(f"📦 Frozen: {getattr(sys, 'frozen', False)}")
        print(f"📂 Current folder structure:")
        print("   ├── core/")
        print("   ├── data/")
        print("   ├── gui/")
        print("   ├── utils/")
        print("   ├── main.py")
        print("   ├── icon.ico          ← For development")
        print("   └── mapping.xlsx")
        print("   📝 Note: icon.ico embedded in .exe when compiled")
        print("")
        print("🔧 XLS SUPPORT SETUP:")
        print("   If you get errors with .xls files, run:")
        print("   pip uninstall xlrd -y")
        print("   pip install xlrd==2.0.1")
        print("")
    
    # Handle command line arguments (for file associations)
    if len(sys.argv) > 1:
        print(f"📄 Command line argument: {sys.argv[1]}")
    
    # Run main application
    main()
