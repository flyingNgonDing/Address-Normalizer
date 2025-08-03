"""
NEW: Sheet selector dialog for Excel files with multiple sheets
FIXED: Added proper flow with "Start Processing" functionality
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys


class SheetSelectorDialog:
    """Dialog để chọn sheets từ Excel file"""
    
    def __init__(self, parent, sheet_names, file_name):
        self.parent = parent
        self.sheet_names = sheet_names
        self.file_name = file_name
        self.selected_sheets = []
        self.result = None
        self.action = None  # NEW: Track what action user took
        
        self.dialog = tk.Toplevel(parent)
        self.setup_dialog()
        
    def setup_dialog(self):
        """Thiết lập dialog"""
        self.dialog.title("Chọn Sheets để xử lý")
        self.dialog.geometry("500x700")  # Increased height for new buttons
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.center_dialog()
        
        # Create UI
        self.create_header()
        self.create_sheet_list()
        self.create_buttons()
        
        # Focus on dialog
        self.dialog.focus_set()
    
    def center_dialog(self):
        """Center dialog on parent window"""
        self.dialog.update_idletasks()
        
        # Get parent window position and size
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # Calculate dialog position
        dialog_width = 500
        dialog_height = 450  # Updated height
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
    def create_header(self):
        """Tạo header của dialog"""
        header_frame = tk.Frame(self.dialog, bg='white', pady=15)
        header_frame.pack(fill='x', padx=20)
        
        title_label = tk.Label(
            header_frame,
            text="Chọn Sheets để xử lý",
            font=('Segoe UI', 14, 'bold') if sys.platform.startswith('win') else ('Arial', 14, 'bold'),
            bg='white',
            fg='#1a1a1a'
        )
        title_label.pack()
        
        file_label = tk.Label(
            header_frame,
            text=f"File: {self.file_name}",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='white',
            fg='#4a4a4a'
        )
        file_label.pack(pady=(5, 0))
        
        instruction_label = tk.Label(
            header_frame,
            text="Chọn các sheet cần xử lý và sắp xếp theo thứ tự mong muốn:",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='white',
            fg='#4a4a4a'
        )
        instruction_label.pack(pady=(10, 0))
    
    def create_sheet_list(self):
        """Tạo danh sách sheets"""
        list_frame = tk.Frame(self.dialog, bg='white')
        list_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # Available sheets section
        available_label = tk.Label(
            list_frame,
            text="Sheets có sẵn:",
            font=('Segoe UI', 11, 'bold') if sys.platform.startswith('win') else ('Arial', 11, 'bold'),
            bg='white',
            fg='#1a1a1a'
        )
        available_label.pack(anchor='w')
        
        # Available sheets listbox with scrollbar
        available_frame = tk.Frame(list_frame, bg='white')
        available_frame.pack(fill='x', pady=(5, 15))
        
        self.available_listbox = tk.Listbox(
            available_frame,
            height=5,  # Reduced height to make room for buttons
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            selectmode=tk.SINGLE
        )
        available_scrollbar = ttk.Scrollbar(available_frame, orient='vertical', command=self.available_listbox.yview)
        self.available_listbox.configure(yscrollcommand=available_scrollbar.set)
        
        self.available_listbox.pack(side='left', fill='both', expand=True)
        available_scrollbar.pack(side='right', fill='y')
        
        # Populate available sheets
        for sheet in self.sheet_names:
            self.available_listbox.insert(tk.END, sheet)
        
        # Control buttons
        control_frame = tk.Frame(list_frame, bg='white')
        control_frame.pack(fill='x', pady=(0, 15))
        
        add_button = tk.Button(
            control_frame,
            text="➤ Thêm",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#1976d2',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            command=self.add_sheet
        )
        add_button.pack(side='left', padx=(0, 10))
        
        remove_button = tk.Button(
            control_frame,
            text="➤ Bỏ",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#d32f2f',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            command=self.remove_sheet
        )
        remove_button.pack(side='left', padx=(0, 10))
        
        up_button = tk.Button(
            control_frame,
            text="↑ Lên",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#2e7d32',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            command=self.move_up
        )
        up_button.pack(side='left', padx=(0, 10))
        
        down_button = tk.Button(
            control_frame,
            text="↓ Xuống",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#2e7d32',
            fg='white',
            relief='flat',
            padx=15,
            pady=5,
            command=self.move_down
        )
        down_button.pack(side='left')
        
        # Selected sheets section
        selected_label = tk.Label(
            list_frame,
            text="Sheets được chọn (theo thứ tự xử lý):",
            font=('Segoe UI', 11, 'bold') if sys.platform.startswith('win') else ('Arial', 11, 'bold'),
            bg='white',
            fg='#1a1a1a'
        )
        selected_label.pack(anchor='w')
        
        # Selected sheets listbox with scrollbar
        selected_frame = tk.Frame(list_frame, bg='white')
        selected_frame.pack(fill='both', expand=True, pady=(5, 0))
        
        self.selected_listbox = tk.Listbox(
            selected_frame,
            height=5,  # Reduced height to make room for buttons
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            selectmode=tk.SINGLE
        )
        selected_scrollbar = ttk.Scrollbar(selected_frame, orient='vertical', command=self.selected_listbox.yview)
        self.selected_listbox.configure(yscrollcommand=selected_scrollbar.set)
        
        self.selected_listbox.pack(side='left', fill='both', expand=True)
        selected_scrollbar.pack(side='right', fill='y')
        
        # Bind double-click events
        self.available_listbox.bind('<Double-Button-1>', lambda e: self.add_sheet())
        self.selected_listbox.bind('<Double-Button-1>', lambda e: self.remove_sheet())
    
    def create_buttons(self):
        """Tạo buttons của dialog - UPDATED with new flow"""
        button_frame = tk.Frame(self.dialog, bg='white', pady=15)
        button_frame.pack(fill='x', padx=20)
        
        # Quick action buttons at top
        quick_frame = tk.Frame(button_frame, bg='white')
        quick_frame.pack(fill='x', pady=(0, 15))
        
        select_all_button = tk.Button(
            quick_frame,
            text="📋 Chọn tất cả",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#2e7d32',
            fg='white',
            relief='flat',
            padx=20,
            pady=8,
            command=self.select_all
        )
        select_all_button.pack(side='left', padx=(0, 10))
        
        # Tìm dòng 227-239 và thay thế:
        first_only_button = tk.Button(
            quick_frame,
            text="🚀 BẮT ĐẦU XỬ LÝ",  # ← Đổi text
            font=('Segoe UI', 12, 'bold'),  # ← Font lớn hơn
            bg='#1976d2',  # ← Màu xanh dương
            fg='white',
            relief='raised',  # ← Nổi bật hơn
            borderwidth=3,
            padx=30,
            pady=12,
            command=self.start_processing  # ← Đổi command
        )
        first_only_button.pack(side='left')
        
        # Main action buttons at bottom
        action_frame = tk.Frame(button_frame, bg='white')
        action_frame.pack(fill='x')
        
        # Cancel button
        cancel_button = tk.Button(
            action_frame,
            text="❌ Hủy",
            font=('Segoe UI', 11) if sys.platform.startswith('win') else ('Arial', 11),
            bg='#d32f2f',
            fg='white',
            relief='flat',
            padx=25,
            pady=10,
            command=self.cancel
        )
        cancel_button.pack(side='right', padx=(10, 0))
        
        # NEW: Start Processing button - MAIN ACTION
        start_button = tk.Button(
            action_frame,
            text="🚀 BẮT ĐẦU XỬ LÝ",
            font=('Segoe UI', 12, 'bold') if sys.platform.startswith('win') else ('Arial', 12, 'bold'),
            bg='#1976d2',
            fg='white',
            relief='raised',
            borderwidth=3,
            padx=30,
            pady=12,
            cursor='hand2',
            command=self.start_processing
        )
        start_button.pack(side='right', padx=(0, 10))
                
        # Preview button
        preview_button = tk.Button(
            action_frame,
            text="👁️ Xem trước",
            font=('Segoe UI', 10) if sys.platform.startswith('win') else ('Arial', 10),
            bg='#757575',
            fg='white',
            relief='flat',
            padx=20,
            pady=8,
            command=self.preview_selection
        )
        preview_button.pack(side='left')
    
    def add_sheet(self):
        """Thêm sheet vào danh sách được chọn"""
        selection = self.available_listbox.curselection()
        if selection:
            index = selection[0]
            sheet_name = self.available_listbox.get(index)
            
            # Add to selected if not already there
            if sheet_name not in [self.selected_listbox.get(i) for i in range(self.selected_listbox.size())]:
                self.selected_listbox.insert(tk.END, sheet_name)
    
    def remove_sheet(self):
        """Bỏ sheet khỏi danh sách được chọn"""
        selection = self.selected_listbox.curselection()
        if selection:
            self.selected_listbox.delete(selection[0])
    
    def move_up(self):
        """Di chuyển sheet lên trên"""
        selection = self.selected_listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            sheet_name = self.selected_listbox.get(index)
            self.selected_listbox.delete(index)
            self.selected_listbox.insert(index - 1, sheet_name)
            self.selected_listbox.selection_set(index - 1)
    
    def move_down(self):
        """Di chuyển sheet xuống dưới"""
        selection = self.selected_listbox.curselection()
        if selection and selection[0] < self.selected_listbox.size() - 1:
            index = selection[0]
            sheet_name = self.selected_listbox.get(index)
            self.selected_listbox.delete(index)
            self.selected_listbox.insert(index + 1, sheet_name)
            self.selected_listbox.selection_set(index + 1)
    
    def select_all(self):
        """Chọn tất cả sheets"""
        self.selected_listbox.delete(0, tk.END)
        for sheet in self.sheet_names:
            self.selected_listbox.insert(tk.END, sheet)
    
    def select_first_only(self):
        """NEW: Chọn chỉ sheet đầu tiên"""
        self.selected_listbox.delete(0, tk.END)
        if self.sheet_names:
            self.selected_listbox.insert(tk.END, self.sheet_names[0])
    
    def preview_selection(self):
        """NEW: Xem trước lựa chọn"""
        selected_sheets = [self.selected_listbox.get(i) for i in range(self.selected_listbox.size())]
        
        if not selected_sheets:
            messagebox.showwarning("Cảnh báo", "Chưa chọn sheet nào!")
            return
        
        preview_text = f"Sẽ xử lý {len(selected_sheets)} sheet(s) theo thứ tự:\n\n"
        for i, sheet in enumerate(selected_sheets, 1):
            preview_text += f"{i}. {sheet}\n"
        
        preview_text += f"\nBạn có muốn tiếp tục?"
        
        messagebox.showinfo("Xem trước", preview_text)
    
    def start_processing(self):
        """Bắt đầu xử lý - MAIN ACTION"""
        selected_sheets = [self.selected_listbox.get(i) for i in range(self.selected_listbox.size())]
        
        if not selected_sheets:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một sheet!")
            return
        
        # Confirm before processing
        confirm_text = (
            f"Bắt đầu xử lý {len(selected_sheets)} sheet(s)?\n\n"
            + "\n".join(f"• {sheet}" for sheet in selected_sheets) +
            f"\n\nQuá trình này có thể mất vài phút."
        )
        
        result = messagebox.askyesno("Xác nhận xử lý", confirm_text, icon='question')
        if not result:
            return
        
        self.result = selected_sheets
        self.action = "start_processing"
        self.dialog.destroy()

    
    def cancel(self):
        """Hủy dialog"""
        self.result = None
        self.action = "cancel"
        self.dialog.destroy()
    
    def show(self):
        """Hiển thị dialog và trả về kết quả"""
        self.dialog.wait_window()
        return {
            'sheets': self.result,
            'action': self.action
        }


def show_sheet_selector(parent, sheet_names, file_name):
    """
    Hiển thị dialog chọn sheets - UPDATED to return action info
    
    Args:
        parent: Parent window
        sheet_names: List of sheet names
        file_name: File name for display
        
    Returns:
        dict: {'sheets': [selected_sheets], 'action': 'start_processing'|'cancel'}
    """
    dialog = SheetSelectorDialog(parent, sheet_names, file_name)
    return dialog.show()