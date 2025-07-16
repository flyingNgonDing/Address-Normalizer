"""
File Processing Logic
Handles file selection, processing, and multi-sheet operations
FIXED: Enhanced .xls support with proper file dialog and error handling
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import time
import os
import sys
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

from config import EXPANDED_GEOMETRY, COLORS, CHUNK_SIZE, FILE_DIALOG_FILETYPES

# Safe import for sheet_selector
try:
    from gui.sheet_selector import show_sheet_selector
except ImportError:
    # Fallback function if sheet_selector is not available
    def show_sheet_selector(parent, sheet_names, file_name):
        """Fallback sheet selector using simple dialog"""
        if len(sheet_names) == 1:
            return {'action': 'start_processing', 'sheets': sheet_names}
        
        # Simple selection dialog
        selection = messagebox.askyesno(
            "Multi-sheet detected",
            f"File có {len(sheet_names)} sheets:\n" + "\n".join(f"- {name}" for name in sheet_names) +
            f"\n\nChọn YES để xử lý tất cả, NO để chỉ xử lý sheet đầu tiên."
        )
        
        if selection:
            return {'action': 'start_processing', 'sheets': sheet_names}  # Process all sheets
        else:
            return {'action': 'start_processing', 'sheets': [sheet_names[0]]}  # Process only first sheet

from core.file_handler import read_file, save_file, save_multiple_sheets, check_required_columns, get_excel_sheet_names
from core.text_processor import chuan_hoa, find_ap_column, find_address_column
from core.fuzzy_matcher import fuzzy_match_row
from utils.performance import detect_mode
from utils.helpers import format_time, format_number


class FileProcessor:
    """Handles all file processing operations"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.root = main_window.root
        self.components = None  # Will be set after components are created
    
    def set_components(self, components):
        """Set reference to window components"""
        self.components = components
    
    def chon_file(self):
        """Chọn file thông qua dialog - FIXED with proper .xls support"""
        try:
            # FIXED: Use proper file dialog format from config
            file_benh_nhan = filedialog.askopenfilename(
                title="Chọn file danh sách bệnh nhân",
                filetypes=FILE_DIALOG_FILETYPES,
                initialdir=self._get_initial_dir()
            )
            
            if not file_benh_nhan:
                return

            # ADDED: Validate file extension
            file_ext = os.path.splitext(file_benh_nhan.lower())[1]
            if file_ext not in ['.xlsx', '.xls', '.csv']:
                messagebox.showerror(
                    "Lỗi định dạng file", 
                    f"File không được hỗ trợ: {file_ext}\n\n"
                    f"Chỉ hỗ trợ: .xlsx, .xls, .csv"
                )
                return

            # ADDED: Check if file can be read
            if not self._validate_file_accessibility(file_benh_nhan):
                return

            self.start_processing(file_benh_nhan)
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi chọn file:\n{str(e)}")
    
    def _get_initial_dir(self):
        """Get initial directory for file dialog"""
        if sys.platform.startswith('win'):
            return os.path.expanduser("~\\Desktop")  # Windows Desktop
        else:
            return os.path.expanduser("~/Desktop")   # Unix Desktop
    
    def _validate_file_accessibility(self, file_path):
        """Validate if file can be accessed and read"""
        try:
            # Check if file exists and is readable
            if not os.path.exists(file_path):
                messagebox.showerror("Lỗi", f"File không tồn tại:\n{file_path}")
                return False
            
            if not os.access(file_path, os.R_OK):
                messagebox.showerror("Lỗi", f"Không có quyền đọc file:\n{file_path}")
                return False
            
            # Check file size (not too large)
            file_size = os.path.getsize(file_path)
            if file_size > 100 * 1024 * 1024:  # 100MB limit
                result = messagebox.askyesno(
                    "Cảnh báo", 
                    f"File rất lớn ({file_size // (1024*1024)} MB).\n"
                    f"Quá trình xử lý có thể mất nhiều thời gian.\n\n"
                    f"Bạn có muốn tiếp tục không?",
                    icon='warning'
                )
                if not result:
                    return False
            
            # ADDED: Quick test to see if file can be opened
            file_ext = os.path.splitext(file_path.lower())[1]
            
            if file_ext in ['.xlsx', '.xls']:
                # Test Excel file reading
                try:
                    get_excel_sheet_names(file_path)
                except Exception as e:
                    error_msg = f"Lỗi đọc file Excel:\n{str(e)}"
                    
                    if file_ext == '.xls':
                        error_msg += f"\n\n💡 GỢI Ý CHO FILE .XLS:"
                        error_msg += f"\n- Đảm bảo đã cài đặt: pip install xlrd==2.0.1"
                        error_msg += f"\n- Thử mở file bằng Excel và lưu lại thành .xlsx"
                        error_msg += f"\n- Kiểm tra file có bị hỏng không"
                    
                    messagebox.showerror("Lỗi đọc file", error_msg)
                    return False
            
            elif file_ext == '.csv':
                # Test CSV file reading
                try:
                    pd.read_csv(file_path, nrows=1)  # Test read first row
                except Exception as e:
                    messagebox.showerror("Lỗi đọc CSV", f"Lỗi đọc file CSV:\n{str(e)}")
                    return False
            
            return True
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi kiểm tra file:\n{str(e)}")
            return False
    
    def process_file_from_path(self, file_path):
        """Xử lý file từ đường dẫn"""
        if not os.path.exists(file_path):
            messagebox.showerror("Lỗi", f"File không tồn tại:\n{file_path}")
            return
        
        # ADDED: Validate file before processing
        if not self._validate_file_accessibility(file_path):
            return
        
        # Hiển thị thông báo xác nhận
        file_name = os.path.basename(file_path)
        result = messagebox.askyesno(
            "Xác nhận", 
            f"Bạn có muốn xử lý file:\n{file_name}?",
            icon='question'
        )
        if not result:
            return
        
        self.start_processing(file_path)
    
    def start_processing(self, file_path):
        """Bắt đầu quá trình xử lý file - ENHANCED with better error handling"""
        try:
            # Check if Excel file and get sheets
            file_ext = os.path.splitext(file_path.lower())[1]
            selected_sheets = None
            
            if file_ext in ['.xlsx', '.xls']:
                try:
                    sheet_names = get_excel_sheet_names(file_path)
                    
                    if len(sheet_names) > 1:
                        # Show sheet selector dialog
                        file_name = os.path.basename(file_path)
                        dialog_result = show_sheet_selector(self.root, sheet_names, file_name)
                        
                        # Handle dialog result
                        if not dialog_result or dialog_result['action'] == 'cancel':
                            return  # User cancelled
                        
                        if dialog_result['action'] == 'start_processing':
                            selected_sheets = dialog_result['sheets']
                            if not selected_sheets:
                                messagebox.showwarning("Cảnh báo", "Không có sheet nào được chọn!")
                                return
                        else:
                            return  # Unknown action
                    else:
                        selected_sheets = [sheet_names[0]]  # Single sheet
                        
                except Exception as e:
                    error_msg = f"Lỗi đọc file Excel:\n{str(e)}"
                    
                    if file_ext == '.xls':
                        error_msg += f"\n\n🔧 CÁCH KHẮC PHỤC FILE .XLS:"
                        error_msg += f"\n1. Cài đặt đúng version xlrd:"
                        error_msg += f"\n   pip uninstall xlrd -y"
                        error_msg += f"\n   pip install xlrd==2.0.1"
                        error_msg += f"\n\n2. Hoặc mở file bằng Excel và lưu thành .xlsx"
                        error_msg += f"\n\n3. Kiểm tra file có bị hỏng không"
                    
                    messagebox.showerror("Lỗi đọc Excel", error_msg)
                    return
            else:
                selected_sheets = [None]  # CSV file
            
            # Continue with existing processing logic...
            # Animate window resize for Windows
            if self.main_window.use_animations:
                self.main_window.animate_window_resize()
            else:
                self.root.geometry(EXPANDED_GEOMETRY)
                
            self.main_window.components.label.config(text="Đang khởi tạo xử lý...")
            
            # Ẩn các element không cần thiết
            self.main_window.components.main_button_frame.pack_forget()
            self.main_window.components.settings_container.pack_forget()
            
            self.main_window.components.control_frame.pack(pady=15)
            self.main_window.components.log_frame.pack(fill="both", expand=True, pady=(0, 10))

            # Reset state
            self.main_window.processing = True
            self.main_window.stop_flag = False
            self.main_window.paused = False
            self.main_window.done_rows = 0
            self.main_window.total_paused_time = 0
            self.main_window.pause_start_time = 0
            self.main_window.start_time = time.time()
            
            # Multi-sheet processing setup
            self.main_window.current_sheet_index = 0
            self.main_window.total_sheets = len(selected_sheets)
            self.main_window.sheet_results = {}
            
            self.update_timer()
            threading.Thread(target=self.xu_ly_file_sheets, args=(file_path, selected_sheets), daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Lỗi khởi tạo", f"Lỗi khởi tạo xử lý:\n{str(e)}")

    def xu_ly_file_sheets(self, file_path, selected_sheets):
        """Xử lý multiple sheets"""
        try:
            self.root.after(0, lambda: self.main_window.components.label.config(
                text=f"Đang xử lý {len(selected_sheets)} sheet(s)..."
            ))
            
            for i, sheet_name in enumerate(selected_sheets):
                if self.main_window.stop_flag:
                    break
                
                self.main_window.current_sheet_index = i + 1
                
                # Update progress for current sheet
                self.root.after(0, lambda s=sheet_name or "Sheet1", idx=i+1: self.main_window.components.label.config(
                    text=f"Đang xử lý sheet {idx}/{len(selected_sheets)}: {s}"
                ))
                
                # Process individual sheet
                result_df = self.xu_ly_single_sheet(file_path, sheet_name, i)
                
                if result_df is not None and not self.main_window.stop_flag:
                    # Store result with proper sheet name
                    result_sheet_name = f"Sheet{i+1}"
                    self.main_window.sheet_results[result_sheet_name] = result_df
                    
                    self.root.after(0, lambda s=sheet_name or "Sheet1": self.main_window.components.update_sheet_log(
                        f"✅ Hoàn thành sheet: {s}"
                    ))
                elif self.main_window.stop_flag:
                    break
                else:
                    self.root.after(0, lambda s=sheet_name or "Sheet1": self.main_window.components.update_sheet_log(
                        f"❌ Lỗi xử lý sheet: {s}"
                    ))
            
            if not self.main_window.stop_flag and self.main_window.sheet_results:
                # Save all results
                self._save_multiple_sheets_result(file_path)
            
        except Exception as e:
            if not self.main_window.stop_flag:
                error_msg = f"Có lỗi xảy ra trong quá trình xử lý:\n{str(e)}"
                self.root.after(0, lambda: messagebox.showerror("Lỗi", error_msg))
                self.root.after(0, lambda: self.main_window.components.label.config(text="❌ Xử lý thất bại"))
        finally:
            self.root.after(0, self.main_window.reset_ui)
    
    def xu_ly_single_sheet(self, file_path, sheet_name, sheet_index):
        """Xử lý một sheet đơn lẻ - ENHANCED with better error handling"""
        try:
            # Đọc data từ sheet
            if sheet_name is not None:
                df = read_file(file_path, sheet_name)
            else:
                df = read_file(file_path)  # CSV
            
            # Kiểm tra dữ liệu rỗng
            if df.empty:
                self.root.after(0, lambda: self.main_window.components.update_sheet_log(
                    f"⚠️ Sheet rỗng: {sheet_name or 'Sheet1'}"
                ))
                return None
            
            # Kiểm tra các cột bắt buộc
            column_check = check_required_columns(df)
            if not column_check['valid']:
                missing_cols = ', '.join(column_check['missing'])
                self.root.after(0, lambda: self.main_window.components.update_sheet_log(
                    f"❌ Sheet thiếu cột: {missing_cols}"
                ))
                return None
            
            xa_col = column_check['xa_col']
            huyen_col = column_check['huyen_col'] 
            tinh_col = column_check['tinh_col']
            
            # Find ấp and address columns
            ap_col = find_ap_column(df)
            address_col = find_address_column(df)
            
            sheet_rows = len(df)
            self.main_window.total_rows = sheet_rows * self.main_window.total_sheets  # Approximate for progress
            
            # Xử lý theo chunks với Windows optimization
            processed_df = self._process_dataframe_chunks_with_ap(
                df, xa_col, huyen_col, tinh_col, ap_col, address_col, sheet_index
            )
            
            return processed_df
            
        except Exception as e:
            error_msg = f"Error processing sheet {sheet_name}: {e}"
            print(error_msg)
            self.root.after(0, lambda: self.main_window.components.update_sheet_log(
                f"❌ Lỗi xử lý sheet {sheet_name}: {str(e)}"
            ))
            return None
    
    def _process_dataframe_chunks_with_ap(self, df, xa_col, huyen_col, tinh_col, ap_col, address_col, sheet_index):
        """Process DataFrame with ấp support"""
        # Use smaller chunks on Windows for better responsiveness
        chunk_size = min(CHUNK_SIZE, 300)
        sheet_rows = len(df)
        n_chunks = (sheet_rows + chunk_size - 1) // chunk_size
        results = [None] * n_chunks
        mode = detect_mode()

        def worker(chunk_index):
            if self.main_window.stop_flag:
                return
            start = chunk_index * chunk_size
            end = min((chunk_index + 1) * chunk_size, sheet_rows)
            chunk = df.iloc[start:end].copy()
            processed = self.process_chunk_with_ap(
                chunk, xa_col, huyen_col, tinh_col, ap_col, address_col, 
                chunk_index, n_chunks, sheet_index
            )
            if not self.main_window.stop_flag:
                results[chunk_index] = processed

        # Use more conservative threading on Windows
        if mode == "serial" or sys.platform.startswith('win'):
            for i in range(n_chunks):
                if self.main_window.stop_flag:
                    break
                worker(i)
        else:
            max_workers = min(2, os.cpu_count() or 1)  # Limit to 2 workers on Windows
            self.main_window.executor = ThreadPoolExecutor(max_workers=max_workers)
            futures = []
            
            for i in range(n_chunks):
                if self.main_window.stop_flag:
                    break
                future = self.main_window.executor.submit(worker, i)
                futures.append(future)

            # Wait for completion with timeout
            timeout_counter = 0
            while not self.main_window.stop_flag and any(r is None for r in results) and timeout_counter < 300:  # 5 minute timeout
                time.sleep(1)
                timeout_counter += 1

            if self.main_window.executor:
                self.main_window.executor.shutdown(wait=False)

        # Gộp kết quả
        valid_results = [r for r in results if r is not None]
        if not valid_results:
            return None

        return pd.concat(valid_results, ignore_index=True)
    
    def process_chunk_with_ap(self, chunk, xa_col, huyen_col, tinh_col, ap_col, address_col, chunk_index=0, chunk_total=1, sheet_index=0):
        """Xử lý một chunk dữ liệu với ấp support - FIXED"""
        # Tạo các cột tạm thời để xử lý
        chunk['_xa_chuan'] = chunk[xa_col].apply(chuan_hoa)
        chunk['_huyen_chuan'] = chunk[huyen_col].apply(chuan_hoa)
        chunk['_tinh_chuan'] = chunk[tinh_col].apply(chuan_hoa)

        chunk['Lý do không match'] = chunk.apply(
            lambda row: 'Thiếu xã/tỉnh' if not row['_xa_chuan'].strip() or not row['_tinh_chuan'].strip() else '',
            axis=1
        )

        results = []
        # FIXED: Sử dụng iterrows() thay vì itertuples() để tránh lỗi attribute
        for idx, row in chunk.iterrows():
            # Kiểm tra pause/stop with shorter sleep for Windows responsiveness
            while self.main_window.paused and not self.main_window.stop_flag:
                time.sleep(0.05)
            if self.main_window.stop_flag:
                return chunk

            if row['Lý do không match']:  # Nếu đã có lý do không match
                results.append((None, None, None, None, None, row['Lý do không match']))
            else:
                # FIXED: Get ấp and address info using proper column access
                ap_value = row[ap_col] if ap_col and ap_col in chunk.columns else None
                address_value = row[address_col] if address_col and address_col in chunk.columns else None
                
                # FIXED: Call fuzzy match with proper row access
                matched = fuzzy_match_row(
                    row['_xa_chuan'], row['_huyen_chuan'], row['_tinh_chuan'],
                    ap=ap_value, address_detail=address_value
                )
                results.append(matched)

            with self.main_window.lock:
                self.main_window.done_rows += 1
            
            # Update log more frequently on Windows for better user feedback
            local_idx = len(results)
            if local_idx % 5 == 0 or local_idx == len(chunk):
                self.main_window.components.update_log_with_sheet(chunk_index, chunk_total, local_idx, len(chunk), sheet_index)

        # Lưu kết quả
        chunk[['_xacu', '_huyencu', '_tinhcu', 'Xã sau sáp nhập', 'Tỉnh sau sáp nhập', 'Lý do không match']] = pd.DataFrame(results, index=chunk.index)
        
        # Xóa các cột tạm thời
        chunk = chunk.drop(columns=['_xa_chuan', '_huyen_chuan', '_tinh_chuan', '_xacu', '_huyencu', '_tinhcu'])
        
        return chunk
    
    def _save_multiple_sheets_result(self, original_file_path):
        """Lưu kết quả multiple sheets"""
        try:
            # FIXED: Windows specific file dialog
            file_luu = filedialog.asksaveasfilename(
                title="Lưu file kết quả",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialdir=self._get_initial_dir()
            )
            
            if not file_luu:
                self.root.after(0, self.main_window.reset_ui)
                return

            # Prepare sheet data for saving (keep original columns + new columns)
            sheets_to_save = {}
            total_processed = 0
            
            for sheet_name, df_result in self.main_window.sheet_results.items():
                # Keep all original columns plus new ones
                sheets_to_save[sheet_name] = df_result
                total_processed += len(df_result)
            
            # Save multiple sheets
            save_multiple_sheets(sheets_to_save, file_luu)
            
            self.root.after(0, lambda: self.main_window.components.label.config(text="✅ Xử lý hoàn tất thành công!"))
            self.root.after(0, lambda: messagebox.showinfo(
                "Hoàn tất", 
                f"Đã xử lý thành công {len(self.main_window.sheet_results)} sheet(s)!\n"
                f"Tổng cộng {format_number(total_processed)} bản ghi.\n\n"
                f"File kết quả đã được lưu tại:\n{file_luu}"
            ))
        except PermissionError:
            self.root.after(0, lambda: messagebox.showerror(
                "Lỗi", 
                f"Không thể lưu file!\n\n"
                f"File có thể đang được mở trong Excel.\n"
                f"Vui lòng đóng Excel và thử lại."
            ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi lưu file: {str(e)}"))
    
    def update_timer(self):
        """Cập nhật timer"""
        if self.main_window.processing and not self.main_window.stop_flag:
            self.update_progress()
            self.root.after(500, self.update_timer)
    
    def update_progress(self):
        """Cập nhật progress bar và thời gian"""
        try:
            done = self.main_window.done_rows
            progress_percentage = done / self.main_window.total_rows * 100 if self.main_window.total_rows else 0
            self.main_window.components.progress['value'] = progress_percentage
            
            # Tính thời gian thực tế đã trải qua
            current_time = time.time()
            
            if self.main_window.paused:
                elapsed = int(self.main_window.pause_start_time - self.main_window.start_time - self.main_window.total_paused_time)
            else:
                elapsed = int(current_time - self.main_window.start_time - self.main_window.total_paused_time)
            
            elapsed = max(0, elapsed)
            
            # Tính thời gian còn lại
            if done > 0 and elapsed > 0:
                remaining = int(elapsed * (self.main_window.total_rows - done) / done)
            else:
                remaining = 0
            
            # Update label with progress percentage and sheet info
            sheet_info = ""
            if self.main_window.total_sheets > 1:
                sheet_info = f" - Sheet {self.main_window.current_sheet_index}/{self.main_window.total_sheets}"
            
            self.main_window.components.time_label.config(
                text=f"{format_time(elapsed)} / {format_time(remaining)} - "
                     f"thời gian đã xử lý / thời gian còn lại ({progress_percentage:.1f}%){sheet_info}"
            )
        except Exception as e:
            print(f"Progress update error: {e}")
    
    def toggle_pause(self):
        """Toggle pause/resume"""
        current_time = time.time()
        
        if not self.main_window.paused:
            self.main_window.paused = True
            self.main_window.pause_start_time = current_time
            # Đổi thành nút "Tiếp Tục" màu xanh
            self.main_window.components.pause_button.config(
                text="Tiếp Tục", 
                bg=COLORS['success'],
                activebackground="#1b5e20"
            )
        else:
            self.main_window.paused = False
            self.main_window.total_paused_time += current_time - self.main_window.pause_start_time
            self.main_window.pause_start_time = 0
            # Đổi về nút "Tạm Dừng" màu đỏ
            self.main_window.components.pause_button.config(
                text="Tạm Dừng", 
                bg=COLORS['danger'],
                activebackground="#b71c1c"
            )
    
    def cancel_process(self):
        """Huỷ hoàn toàn tiến trình xử lý"""
        result = messagebox.askyesno(
            "Xác nhận", 
            "Bạn có chắc chắn muốn huỷ tiến trình đang xử lý?",
            icon='warning'
        )
        if result:
            self.main_window.stop_flag = True
            if self.main_window.executor:
                self.main_window.executor.shutdown(wait=False)
