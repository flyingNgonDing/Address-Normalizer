"""
Module xử lý drag & drop functionality
WINDOWS VERSION - FIXED: Enable drag & drop for frozen executable
"""
import os
import re
import urllib.parse
import sys

# ✅ FIXED: Always try to import tkinterdnd2, including in frozen executable
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
    print("✅ tkinterdnd2 loaded successfully (including frozen executable)")
except ImportError:
    DND_AVAILABLE = False
    print("⚠️ tkinterdnd2 not available. Using alternative file selection method.")

# Import Windows-specific modules if available
try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# Import constants from config - avoid importing GUI modules here
try:
    from config import COLORS, SUPPORTED_EXTENSIONS
except ImportError:
    # Fallback constants if config not available
    COLORS = {
        'bg_primary': '#f0f0f0',
        'bg_drag': '#e1f5fe',
        'text_primary': '#2c3e50',
        'text_secondary': '#7f8c8d',
    }
    SUPPORTED_EXTENSIONS = ('.xlsx', '.xls', '.csv')


class DragDropHandler:
    """Class xử lý drag & drop functionality với Windows optimization"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.drag_active = False
        self.original_bg_color = COLORS['bg_primary']
        self.drag_bg_color = COLORS['bg_drag']
        self.use_alternative_method = not DND_AVAILABLE
    
    def setup_drag_drop(self):
        """Thiết lập tính năng kéo thả file với Windows fallback"""
        if DND_AVAILABLE:
            return self._setup_tkinterdnd2()
        else:
            return self._setup_alternative_method()
    
    def _setup_tkinterdnd2(self):
        """Setup drag & drop using tkinterdnd2"""
        try:
            # Kích hoạt drag & drop cho cửa sổ chính
            self.main_window.root.drop_target_register(DND_FILES)
            self.main_window.root.dnd_bind('<<DropEnter>>', self.on_drop_enter)
            self.main_window.root.dnd_bind('<<DropLeave>>', self.on_drop_leave)
            self.main_window.root.dnd_bind('<<Drop>>', self.on_drop)
            
            # FIXED: Check if components exist before accessing them
            if hasattr(self.main_window, 'components') and self.main_window.components:
                # Kích hoạt cho các widget con quan trọng
                widgets_to_register = []
                
                # Only add widgets that exist
                if hasattr(self.main_window.components, 'main_container') and self.main_window.components.main_container:
                    widgets_to_register.append(self.main_window.components.main_container)
                
                if hasattr(self.main_window.components, 'header_frame') and self.main_window.components.header_frame:
                    widgets_to_register.append(self.main_window.components.header_frame)
                
                if hasattr(self.main_window.components, 'main_button_frame') and self.main_window.components.main_button_frame:
                    widgets_to_register.append(self.main_window.components.main_button_frame)
                
                for widget in widgets_to_register:
                    try:
                        if hasattr(widget, 'drop_target_register'):
                            widget.drop_target_register(DND_FILES)
                            widget.dnd_bind('<<DropEnter>>', self.on_drop_enter)
                            widget.dnd_bind('<<DropLeave>>', self.on_drop_leave)
                            widget.dnd_bind('<<Drop>>', self.on_drop)
                    except Exception as e:
                        print(f"Warning: Could not register drag&drop for widget: {e}")
            
            print("✅ Drag & drop setup successful with tkinterdnd2")
            return True
            
        except Exception as e:
            print(f"❌ Failed to setup tkinterdnd2: {e}")
            return self._setup_alternative_method()
    
    def _setup_alternative_method(self):
        """Setup alternative method without drag & drop"""
        try:
            # Update UI to indicate drag & drop is not available
            if (hasattr(self.main_window, 'components') and 
                self.main_window.components and 
                hasattr(self.main_window.components, 'drag_hint') and 
                self.main_window.components.drag_hint):
                self.main_window.components.drag_hint.configure(
                    text="Sử dụng nút 'Chọn file' để tải file lên"
                )
            
            # Setup keyboard shortcuts as alternative
            self.main_window.root.bind('<Control-o>', lambda e: self.main_window.chon_file())
            
            print("⚠️ Using alternative file selection method")
            return False
            
        except Exception as e:
            print(f"❌ Error setting up alternative method: {e}")
            return False
    
    def on_drop_enter(self, event):
        """Xử lý khi file được kéo vào cửa sổ - Windows optimized"""
        if not self.main_window.processing and not self.drag_active:
            self.drag_active = True
            self.update_drag_visual(True)
    
    def on_drop_leave(self, event):
        """Xử lý khi file được kéo ra khỏi cửa sổ - Windows optimized"""
        if self.drag_active:
            # Check if mouse is really outside the window
            try:
                x, y = self.main_window.root.winfo_pointerxy()
                win_x = self.main_window.root.winfo_rootx()
                win_y = self.main_window.root.winfo_rooty()
                win_width = self.main_window.root.winfo_width()
                win_height = self.main_window.root.winfo_height()
                
                if not (win_x <= x <= win_x + win_width and win_y <= y <= win_y + win_height):
                    self.drag_active = False
                    self.update_drag_visual(False)
            except:
                # Fallback: always reset on leave
                self.drag_active = False
                self.update_drag_visual(False)
    
    def on_drop(self, event):
        """Xử lý khi file được thả vào cửa sổ - Windows optimized"""
        if self.main_window.processing:
            return
        
        self.drag_active = False
        self.update_drag_visual(False)
        
        # Parse file paths từ event data với Windows optimization
        file_paths = self._parse_dropped_files_windows(event.data)
        
        print(f"🔍 Drag & Drop - All detected paths: {file_paths}")
        
        # Kiểm tra số lượng file
        if len(file_paths) == 0:
            self._show_error("Không thể xác định file được kéo thả!")
            print(f"❌ No valid file paths found from: {repr(event.data)}")
            return
        elif len(file_paths) > 1:
            existing_files = [path for path in file_paths if os.path.exists(path)]
            print(f"🔍 Existing files: {existing_files}")
            
            if len(existing_files) > 1:
                self._show_multiple_files_warning(existing_files)
                return
            elif len(existing_files) == 1:
                file_paths = existing_files
        
        # Xử lý file đầu tiên
        if file_paths:
            file_path = file_paths[0]
            self._process_single_file(file_path, event.data)
    
    def _show_error(self, message):
        """Show error message - avoid importing tkinter.messagebox in module level"""
        try:
            from tkinter import messagebox
            messagebox.showerror("Lỗi", message)
        except:
            print(f"❌ Error: {message}")
    
    def _show_warning(self, title, message):
        """Show warning message"""
        try:
            from tkinter import messagebox
            messagebox.showwarning(title, message)
        except:
            print(f"⚠️ Warning - {title}: {message}")
    
    def _parse_dropped_files_windows(self, files_data):
        """Parse file paths từ event data - Windows optimized"""
        print(f"🔍 Raw event data: {repr(files_data)}")
        
        file_paths = []
        
        if isinstance(files_data, (list, tuple)):
            file_paths = [str(f) for f in files_data]
        elif isinstance(files_data, str):
            file_paths = self._parse_string_data_windows(files_data)
        else:
            file_paths = [str(files_data)]
        
        # Làm sạch và chuẩn hóa file paths cho Windows
        cleaned_paths = []
        for path in file_paths:
            if path:
                # Windows specific path cleaning
                path = path.strip().strip('{}').strip('"').strip("'")
                
                # Handle Windows path separators
                path = path.replace('/', '\\')
                
                # Remove file:// protocol if present
                if path.startswith('file://'):
                    path = path[7:]
                
                # URL decode
                try:
                    path = urllib.parse.unquote(path)
                except:
                    pass
                
                if path and os.path.exists(path):
                    cleaned_paths.append(os.path.normpath(path))
        
        return cleaned_paths
    
    def _parse_string_data_windows(self, files_data):
        """Parse string data để extract file paths - Windows version"""
        files_data = files_data.strip()
        
        # Windows specific parsing patterns
        patterns = [
            # Pattern 1: Space-separated paths in quotes
            r'"([^"]+)"',
            # Pattern 2: Paths separated by newlines
            r'([A-Za-z]:\\[^\r\n]+)',
            # Pattern 3: Simple space separation
            r'([A-Za-z]:\\[^\s]+)',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, files_data)
            if matches:
                valid_paths = []
                for match in matches:
                    if os.path.exists(match):
                        valid_paths.append(match)
                if valid_paths:
                    return valid_paths
        
        # Fallback: try splitting by common delimiters
        for delimiter in ['\n', '\r\n', '\t', '  ']:
            if delimiter in files_data:
                parts = [p.strip() for p in files_data.split(delimiter) if p.strip()]
                valid_parts = [p for p in parts if os.path.exists(p)]
                if valid_parts:
                    return valid_parts
        
        # Final fallback: return as single file
        return [files_data] if files_data else []
    
    def _show_multiple_files_warning(self, existing_files):
        """Hiển thị cảnh báo khi kéo thả nhiều file"""
        message = (f"Chỉ có thể xử lý 1 file trong 1 lần!\n\n"
                  f"Bạn đã kéo thả {len(existing_files)} file:\n" + 
                  "\n".join([f"• {os.path.basename(f)}" for f in existing_files[:5]]) +
                  (f"\n• ... và {len(existing_files)-5} file khác" if len(existing_files) > 5 else "") +
                  f"\n\nVui lòng kéo thả từng file một.")
        
        self._show_warning("Cảnh báo", message)
    
    def _process_single_file(self, file_path, original_data):
        """Xử lý single file - Windows optimized"""
        print(f"🔍 Final file path: {repr(file_path)}")
        print(f"🔍 File exists: {os.path.exists(file_path)}")
        
        # Windows specific path fixes
        if not os.path.exists(file_path):
            file_path = self._fix_windows_path(file_path)
        
        # Kiểm tra file có tồn tại không
        if not os.path.exists(file_path):
            self._show_error(f"File không tồn tại!\n\n"
                           f"Đường dẫn đã thử:\n{file_path}\n\n"
                           f"Dữ liệu gốc:\n{repr(original_data)}\n\n"
                           f"Vui lòng thử sử dụng nút 'Chọn file' thay vì kéo thả.")
            return
        
        # Kiểm tra định dạng file
        if not self._validate_file_format(file_path):
            _, file_extension = os.path.splitext(file_path.lower())
            self._show_error(f"File không được hỗ trợ!\n\n"
                           f"File: {os.path.basename(file_path)}\n"
                           f"Extension: {file_extension}\n"
                           f"Định dạng hỗ trợ: {', '.join(SUPPORTED_EXTENSIONS)}")
            return
        
        print(f"✅ Processing file via drag & drop: {file_path}")
        # Call main window method to process file
        if hasattr(self.main_window, 'process_file_from_path'):
            self.main_window.process_file_from_path(file_path)
        else:
            print("❌ main_window doesn't have process_file_from_path method")
    
    def _validate_file_format(self, file_path):
        """Validate file format"""
        if not file_path:
            return False
        
        file_ext = os.path.splitext(file_path.lower())[1]
        return file_ext in SUPPORTED_EXTENSIONS
    
    def _fix_windows_path(self, file_path):
        """Thử sửa file path cho Windows"""
        # Method 1: Normalize path
        normalized = os.path.normpath(file_path)
        if os.path.exists(normalized):
            return normalized
        
        # Method 2: Try with different case (Windows is case-insensitive)
        if os.path.dirname(file_path):
            dir_path = os.path.dirname(file_path)
            filename = os.path.basename(file_path)
            
            if os.path.exists(dir_path):
                # Try to find file with case-insensitive search
                try:
                    for item in os.listdir(dir_path):
                        if item.lower() == filename.lower():
                            return os.path.join(dir_path, item)
                except:
                    pass
        
        # Method 3: Remove extra quotes or spaces
        cleaned = file_path.strip(' "\'')
        if os.path.exists(cleaned):
            return cleaned
        
        return file_path  # Return original if no fix
    
    def update_drag_visual(self, is_dragging):
        """Cập nhật giao diện khi kéo thả file - Windows optimized"""
        try:
            # FIXED: Only update widgets that exist
            if (hasattr(self.main_window, 'components') and 
                self.main_window.components):
                
                if is_dragging:
                    # Thay đổi màu nền khi kéo file vào
                    widgets_to_update = [
                        self.main_window.root,
                    ]
                    
                    # Only add widgets that exist
                    if hasattr(self.main_window.components, 'main_container') and self.main_window.components.main_container:
                        widgets_to_update.append(self.main_window.components.main_container)
                    if hasattr(self.main_window.components, 'header_frame') and self.main_window.components.header_frame:
                        widgets_to_update.append(self.main_window.components.header_frame)
                    if hasattr(self.main_window.components, 'main_button_frame') and self.main_window.components.main_button_frame:
                        widgets_to_update.append(self.main_window.components.main_button_frame)
                    if hasattr(self.main_window.components, 'progress_frame') and self.main_window.components.progress_frame:
                        widgets_to_update.append(self.main_window.components.progress_frame)
                    
                    for widget in widgets_to_update:
                        try:
                            widget.configure(bg=self.drag_bg_color)
                        except:
                            pass
                    
                    # Thay đổi text để hiển thị hướng dẫn
                    if hasattr(self.main_window.components, 'label') and self.main_window.components.label:
                        try:
                            self.main_window.components.label.configure(
                                text="📁 Thả file vào đây để xử lý", 
                                fg='#1976d2', 
                                bg=self.drag_bg_color
                            )
                        except:
                            pass
                    
                    if hasattr(self.main_window.components, 'drag_hint') and self.main_window.components.drag_hint:
                        try:
                            self.main_window.components.drag_hint.configure(
                                text=f"Hỗ trợ file: {', '.join(SUPPORTED_EXTENSIONS)}",
                                fg='#1976d2',
                                bg=self.drag_bg_color
                            )
                        except:
                            pass
                else:
                    # Khôi phục màu nền ban đầu
                    widgets_to_update = [
                        self.main_window.root,
                    ]
                    
                    # Only add widgets that exist
                    if hasattr(self.main_window.components, 'main_container') and self.main_window.components.main_container:
                        widgets_to_update.append(self.main_window.components.main_container)
                    if hasattr(self.main_window.components, 'header_frame') and self.main_window.components.header_frame:
                        widgets_to_update.append(self.main_window.components.header_frame)
                    if hasattr(self.main_window.components, 'main_button_frame') and self.main_window.components.main_button_frame:
                        widgets_to_update.append(self.main_window.components.main_button_frame)
                    if hasattr(self.main_window.components, 'progress_frame') and self.main_window.components.progress_frame:
                        widgets_to_update.append(self.main_window.components.progress_frame)
                    
                    for widget in widgets_to_update:
                        try:
                            widget.configure(bg=self.original_bg_color)
                        except:
                            pass
                    
                    # Khôi phục text ban đầu
                    if hasattr(self.main_window.components, 'label') and self.main_window.components.label:
                        try:
                            self.main_window.components.label.configure(
                                text="Sẵn sàng xử lý danh sách bệnh nhân",
                                fg=COLORS['text_primary'],
                                bg=self.original_bg_color
                            )
                        except:
                            pass
                    
                    if hasattr(self.main_window.components, 'drag_hint') and self.main_window.components.drag_hint:
                        try:
                            hint_text = "Kéo thả file vào đây hoặc nhấn nút bên dưới" if DND_AVAILABLE else "Sử dụng nút 'Chọn file' để tải file lên"
                            self.main_window.components.drag_hint.configure(
                                text=hint_text,
                                fg=COLORS['text_secondary'],
                                bg=self.original_bg_color
                            )
                        except:
                            pass
                        
        except Exception as e:
            print(f"❌ Error updating drag visual: {e}")


def is_drag_drop_available():
    """Kiểm tra xem drag & drop có available không"""
    return DND_AVAILABLE


def get_drag_drop_info():
    """Lấy thông tin về drag & drop support"""
    return {
        'tkinterdnd2_available': DND_AVAILABLE,
        'win32_available': WIN32_AVAILABLE,
        'alternative_method': not DND_AVAILABLE
    }