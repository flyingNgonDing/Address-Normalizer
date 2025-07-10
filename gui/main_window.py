"""
Main GUI window - UPDATED: Removed animation for settings button
"""
import tkinter as tk
from tkinter import messagebox
import threading
import time
import os
import sys

from config import DEFAULT_GEOMETRY, EXPANDED_GEOMETRY, COLORS
from gui.styles import setup_ui_styles, AnimationHelper, apply_windows_theme
from gui.drag_drop import DragDropHandler, is_drag_drop_available
from gui.window_components import WindowComponents
from gui.file_processor import FileProcessor
from gui.animation_handler import AnimationHandler
from utils.performance import ProcessingStats
from utils.helpers import open_file_with_system, get_mapping_file_path


class MainWindow:
    """Main application window - Updated: Simple settings button"""
    
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        
        # Initialize components
        self.components = WindowComponents(self)
        self.file_processor = FileProcessor(self)
        self.animation_handler = AnimationHandler(self)
        
        # Connect components
        self.file_processor.set_components(self.components)
        self.animation_handler.set_components(self.components)
        
        self.setup_ui()
        self.setup_drag_drop()
        self.setup_windows_specific()
        
        # Force update to ensure author info is visible
        self.root.after(200, self.ensure_author_info_visible)
        
    def setup_window(self):
        """Thiết lập cửa sổ chính - Windows optimized"""
        self.root.title("Chuẩn hoá địa chỉ bệnh nhân - Windows")
        self.root.geometry(DEFAULT_GEOMETRY)
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS['bg_primary'])
        
        # Windows specific window settings
        if sys.platform.startswith('win'):
            try:
                # Set window icon if available
                icon_path = os.path.join(os.path.dirname(__file__), '..', 'icon.ico')
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                
                # Set taskbar properties
                self.root.wm_attributes('-toolwindow', False)
                
                # Center window on screen
                self.center_window()
                
            except Exception as e:
                print(f"Warning: Could not set Windows-specific properties: {e}")
    
    def center_window(self):
        """Center window on screen"""
        try:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            pos_x = (self.root.winfo_screenwidth() // 2) - (width // 2)
            pos_y = (self.root.winfo_screenheight() // 2) - (height // 2)
            self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        except:
            pass
    
    def setup_variables(self):
        """Khởi tạo các biến instance"""
        # Processing state
        self.processing = False
        self.paused = False
        self.stop_flag = False
        
        # Threading
        self.lock = threading.Lock()
        self.executor = None
        
        # Statistics
        self.stats = ProcessingStats()
        self.start_time = 0
        self.total_paused_time = 0
        self.pause_start_time = 0
        self.total_rows = 1
        self.done_rows = 0
        
        # Multi-sheet processing
        self.current_sheet_index = 0
        self.total_sheets = 1
        self.sheet_results = {}
        
        # Animation - UPDATED: Disable settings animation
        self.animation_helper = AnimationHelper()
        self.use_animations = self.animation_helper.should_use_animation()
        
        settings = self.animation_helper.get_animation_settings()
        self.animation_steps = settings['steps']
        self.animation_delay = settings['delay']
    
    def setup_styles(self):
        """Thiết lập styles"""
        self.style_manager = setup_ui_styles()
        apply_windows_theme(self.root)
    
    def setup_ui(self):
        """Thiết lập giao diện người dùng"""
        self.components.create_all_components()
    
    def ensure_author_info_visible(self):
        """Ensure author info is visible"""
        try:
            if (hasattr(self.components, 'author_info_frame') and 
                self.components.author_info_frame):
                
                # Make sure it's positioned correctly
                self.components.author_info_frame.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)
                self.components.author_info_frame.lift()
                
                # Check if labels exist and are visible
                if (self.components.author_label1 and 
                    self.components.author_label1.winfo_exists()):
                    print("✅ Author info should now be visible")
                else:
                    print("⚠️ Author labels not found")
            else:
                print("⚠️ Author info frame not found")
                
        except Exception as e:
            print(f"Error ensuring author info visibility: {e}")
    
    def setup_drag_drop(self):
        """Thiết lập drag & drop"""
        self.drag_drop_handler = DragDropHandler(self)
        success = self.drag_drop_handler.setup_drag_drop()
        
        if not success:
            print("Drag & drop not available, using alternative file selection")
    
    def setup_windows_specific(self):
        """Thiết lập các tính năng đặc biệt cho Windows"""
        try:
            # Keyboard shortcuts
            self.root.bind('<Control-o>', lambda e: self.chon_file())
            self.root.bind('<Control-q>', lambda e: self.root.quit())
            self.root.bind('<F1>', lambda e: self.show_help())
            self.root.bind('<Control-s>', lambda e: self.mo_file_mapping())  # NEW: Ctrl+S for settings
            
            # Windows specific event handling
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            # Handle Windows file associations if needed
            if len(sys.argv) > 1:
                file_path = sys.argv[1]
                if os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls', '.csv')):
                    self.root.after(1000, lambda: self.process_file_from_path(file_path))
                    
        except Exception as e:
            print(f"Warning: Could not setup Windows-specific features: {e}")
    
    def on_closing(self):
        """Xử lý khi đóng cửa sổ"""
        if self.processing:
            result = messagebox.askyesno(
                "Xác nhận đóng", 
                "Đang có tiến trình xử lý. Bạn có chắc chắn muốn đóng ứng dụng?",
                icon='warning'
            )
            if not result:
                return
            
            # Stop processing
            self.stop_flag = True
            if self.executor:
                self.executor.shutdown(wait=False)
        
        self.root.quit()
        self.root.destroy()
    
    def show_help(self):
        """Hiển thị help"""
        help_text = """
HƯỚNG DẪN SỬ DỤNG

1. Chọn file danh sách bệnh nhân (Excel hoặc CSV)
2. File phải có các cột: Xã, Huyện, Tỉnh
3. Với Excel nhiều sheet: chọn sheets cần xử lý
4. Chương trình sẽ tự động chuẩn hóa địa chỉ
5. Kết quả được lưu vào file Excel mới

PHÍM TẮT:
- Ctrl+O: Chọn file
- Ctrl+S: Mở cài đặt mapping  ⭐ MỚI
- Ctrl+Q: Thoát
- F1: Hiển thị trợ giúp

CÀI ĐẶT:
- Nút ⚙ xám ở góc dưới phải (phía trên thông tin tác giả)
- Hoặc nhấn Ctrl+S

AUTO-RESTART:
- Khi chỉnh sửa mapping.xlsx và lưu file
- Phần mềm sẽ tự động khởi động lại (KHÔNG có console)
- Nếu không có watchdog: restart thủ công

HỖ TRỢ:
- Windows 8, 10, 11 (32-bit và 64-bit)
- File: .xlsx, .xls, .csv
- Multi-sheet Excel processing
        """
        
        messagebox.showinfo("Trợ giúp", help_text)
    
    def mo_file_mapping(self):
        """Mở file mapping.xlsx để chỉnh sửa"""
        return self.components.mo_file_mapping()
    
    # Delegate methods to components
    def chon_file(self):
        """Chọn file thông qua dialog"""
        return self.file_processor.chon_file()
    
    def process_file_from_path(self, file_path):
        """Xử lý file từ đường dẫn"""
        return self.file_processor.process_file_from_path(file_path)
    
    def toggle_pause(self):
        """Toggle pause/resume"""
        return self.file_processor.toggle_pause()
    
    def cancel_process(self):
        """Huỷ hoàn toàn tiến trình xử lý"""
        return self.file_processor.cancel_process()
    
    def reset_ui(self):
        """Reset UI về trạng thái ban đầu"""
        result = self.components.reset_ui()
        
        # Ensure author info is visible after reset
        self.root.after(100, self.ensure_author_info_visible)
        
        return result
    
    def update_timer(self):
        """Cập nhật timer"""
        return self.file_processor.update_timer()
    
    def update_progress(self):
        """Cập nhật progress bar và thời gian"""
        return self.file_processor.update_progress()
    
    # Animation methods - UPDATED: Simplified for window resize only
    def animate_window_resize(self):
        """Animate window resize for Windows"""
        if self.use_animations:
            result = self.animation_handler.animate_window_resize()
        else:
            self.root.geometry(EXPANDED_GEOMETRY)
            result = None
        
        # Reposition elements for expanded window
        if (hasattr(self.components, 'author_info_frame') and 
            self.components.author_info_frame):
            self.components.author_info_frame.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)
        
        if (hasattr(self.components, 'settings_container') and 
            self.components.settings_container):
            self.components.settings_container.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-45)
        
        return result
    
    def animate_window_resize_back(self):
        """Animate window resize back to original size"""
        if self.use_animations:
            result = self.animation_handler.animate_window_resize_back()
        else:
            self.root.geometry(DEFAULT_GEOMETRY)
            result = None
        
        # Reposition elements for default window
        if (hasattr(self.components, 'author_info_frame') and 
            self.components.author_info_frame):
            self.components.author_info_frame.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)
        
        if (hasattr(self.components, 'settings_container') and 
            self.components.settings_container):
            self.components.settings_container.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-45)
        
        return result