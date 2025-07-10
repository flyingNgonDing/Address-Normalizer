"""
Window UI Components - FIXED author info visibility + Simple settings button
Handles creation and management of UI elements
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys

from config import COLORS, DEFAULT_GEOMETRY
from gui.drag_drop import is_drag_drop_available
from utils.helpers import open_file_with_system, get_mapping_file_path


class WindowComponents:
    """Manages all UI components for the main window"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.root = main_window.root
        self.style_manager = main_window.style_manager
        
        # UI element references
        self.main_container = None
        self.header_frame = None
        self.progress_frame = None
        self.log_frame = None
        self.main_button_frame = None
        self.settings_container = None
        self.control_frame = None
        self.author_info_frame = None
        
        # Individual components
        self.label = None
        self.drag_hint = None
        self.progress = None
        self.time_label = None
        self.log_box = None
        self.button = None
        self.settings_button = None
        self.pause_button = None
        self.cancel_button = None
        self.author_label1 = None
        self.author_label2 = None
    
    def create_all_components(self):
        """Create all UI components"""
        self.create_main_container()
        self.create_header_section()
        self.create_progress_section()
        self.create_log_section()
        self.create_main_button()
        self.create_control_buttons()
        
        # Create author info FIRST, at the very bottom
        self.create_author_info()
        
        # Then create settings button above it
        self.create_settings_button()
    
    def create_main_container(self):
        """Tạo container chính"""
        self.main_container = tk.Frame(self.root, bg=COLORS['bg_primary'])
        self.main_container.pack(fill="both", expand=True, padx=15, pady=15)
        print("✅ Main container created")
    
    def create_header_section(self):
        """Tạo phần header"""
        self.header_frame = tk.Frame(self.main_container, bg=COLORS['bg_primary'])
        self.header_frame.pack(fill="x", pady=(0, 20))

        self.label = tk.Label(
            self.header_frame, 
            text="Sẵn sàng xử lý danh sách bệnh nhân", 
            font=self.style_manager.get_font('custom'),
            bg=COLORS['bg_primary'],
            fg=COLORS['text_primary']
        )
        self.label.pack(pady=(10, 5))

        # Drag & Drop hint with Windows-specific text
        drag_available = is_drag_drop_available()
        hint_text = "Kéo thả file vào đây hoặc nhấn nút bên dưới" if drag_available else "Nhấn nút bên dưới để chọn file"
        
        self.drag_hint = tk.Label(
            self.header_frame,
            text=hint_text,
            font=self.style_manager.get_font('small'),
            bg=COLORS['bg_primary'],
            fg=COLORS['text_secondary']
        )
        self.drag_hint.pack(pady=(0, 5))
    
    def create_progress_section(self):
        """Tạo phần progress"""
        self.progress_frame = tk.Frame(self.main_container, bg=COLORS['bg_primary'])
        self.progress_frame.pack(fill="x", pady=(0, 15))

        self.progress = ttk.Progressbar(
            self.progress_frame, 
            orient='horizontal', 
            length=500, 
            mode='determinate',
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=8)

        self.time_label = tk.Label(
            self.progress_frame, 
            text="00:00 / 00:00 - thời gian đã xử lý / thời gian còn lại", 
            font=self.style_manager.get_font('small'),
            bg=COLORS['bg_primary'],
            fg=COLORS['text_secondary']
        )
        self.time_label.pack()
    
    def create_log_section(self):
        """Tạo phần log"""
        self.log_frame = tk.Frame(self.main_container, bg=COLORS['bg_primary'])
        
        # Create scrolled text widget for Windows
        self.create_scrolled_text()
        self.log_frame.pack_forget()
    
    def create_scrolled_text(self):
        """Tạo text widget với scrollbar cho Windows"""
        # Frame chứa text và scrollbar
        text_frame = tk.Frame(self.log_frame, bg=COLORS['bg_primary'])
        text_frame.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Text widget với contrast cao
        self.log_box = tk.Text(
            text_frame, 
            height=8, 
            font=self.style_manager.get_font('console'),
            bg='#ffffff',
            fg='#1a1a1a',
            relief='solid',
            borderwidth=2,
            highlightthickness=0,
            selectbackground=COLORS['accent'],
            selectforeground='white',
            wrap='word'
        )
        
        # Scrollbar cho Windows
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scrollbar.set)
        
        # Pack với sticky để scrollbar luôn ở bên phải
        self.log_box.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.log_box.config(state='disabled')
    
    def create_main_button(self):
        """Tạo nút chính với màu nền xanh dương đậm"""
        self.main_button_frame = tk.Frame(self.main_container, bg=COLORS['bg_primary'])
        self.main_button_frame.pack(pady=25)

        # Sử dụng tk.Button thay vì ttk.Button để có màu nền
        self.button = tk.Button(
            self.main_button_frame, 
            text="CHỌN FILE DANH SÁCH BỆNH NHÂN",
            font=self.style_manager.get_font('main_button'),
            bg=COLORS['main_button'],      # Màu nền xanh dương đậm
            fg="white",                    # Chữ màu trắng
            activebackground="#1565c0",    # Màu khi hover
            activeforeground="white",      # Màu chữ khi hover
            relief="raised",               # Kiểu nút nổi
            borderwidth=3,                 # Độ dày viền
            padx=30,                       # Padding ngang
            pady=15,                       # Padding dọc
            cursor="hand2",                # Con trỏ tay khi hover
            command=self.main_window.chon_file
        )
        self.button.pack()
        
        print(f"✅ Main button created with background: {COLORS['main_button']}")
        
        # Keyboard shortcut hint for Windows
        shortcut_label = tk.Label(
            self.main_button_frame,
            text="(Ctrl+O)",
            font=self.style_manager.get_font('small'),
            bg=COLORS['bg_primary'],
            fg=COLORS['text_secondary']
        )
        shortcut_label.pack(pady=(8, 0))
    
    def create_control_buttons(self):
        """Tạo các nút control với màu nền thật"""
        self.control_frame = tk.Frame(self.main_container, bg=COLORS['bg_primary'])
        
        # Sử dụng tk.Button để có màu nền thật
        self.pause_button = tk.Button(
            self.control_frame, 
            text="Tạm Dừng",
            font=self.style_manager.get_font('custom'),
            bg=COLORS['danger'],           # Màu đỏ
            fg="white",
            activebackground="#b71c1c",
            activeforeground="white",
            relief="raised",
            borderwidth=2,
            padx=15,
            pady=8,
            cursor="hand2",
            command=self.main_window.toggle_pause
        )
        self.pause_button.pack(side=tk.LEFT, padx=8)

        self.cancel_button = tk.Button(
            self.control_frame, 
            text="Huỷ Tiến Trình",
            font=self.style_manager.get_font('custom'),
            bg=COLORS['accent'],           # Màu xanh
            fg="white",
            activebackground="#0d47a1",
            activeforeground="white",
            relief="raised",
            borderwidth=2,
            padx=15,
            pady=8,
            cursor="hand2",
            command=self.main_window.cancel_process
        )
        self.cancel_button.pack(side=tk.LEFT, padx=8)
    
    def create_author_info(self):
        """FIXED: Tạo thông tin tác giả ở góc dưới phải"""
        print("🎯 Creating author info...")
        
        # Destroy existing if any
        if hasattr(self, 'author_info_frame') and self.author_info_frame:
            try:
                self.author_info_frame.destroy()
            except:
                pass
        
        # Create new frame directly on root (not main_container)
        self.author_info_frame = tk.Frame(self.root, bg=COLORS['bg_primary'])
        
        # Use place instead of pack for absolute positioning
        self.author_info_frame.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)
        
        # Dòng 1: Thông tin viện/khoa/nhóm
        self.author_label1 = tk.Label(
            self.author_info_frame,
            text="Sản phẩm thuộc sở hữu của Viện Pasteur TP HCM - Khoa KSPNBT",
            font=('Segoe UI', 8) if sys.platform.startswith('win') else ('Arial', 8),
            bg=COLORS['bg_primary'],
            fg='#666666',  # Màu xám nhạt
            anchor='e'  # Right-aligned
        )
        self.author_label1.pack(anchor='e')
        
        # Dòng 2: Thông tin tác giả
        self.author_label2 = tk.Label(
            self.author_info_frame,
            text="Liên hệ: Đinh Văn Ngôn (08888 31135)",
            font=('Segoe UI', 8) if sys.platform.startswith('win') else ('Arial', 8),
            bg=COLORS['bg_primary'],
            fg='#666666',  # Màu xám nhạt
            anchor='e'  # Right-aligned
        )
        self.author_label2.pack(anchor='e')
        
        # Force update and bring to front
        self.root.update_idletasks()
        self.author_info_frame.lift()
        
        print("✅ Author info created with place() positioning")
        print(f"🔍 Author frame exists: {bool(self.author_info_frame.winfo_exists())}")
        
        # Debug positioning
        try:
            self.root.after(100, self.debug_author_info_position)
        except:
            pass
    
    def debug_author_info_position(self):
        """Debug author info position"""
        try:
            if self.author_info_frame and self.author_info_frame.winfo_exists():
                x = self.author_info_frame.winfo_x()
                y = self.author_info_frame.winfo_y()
                width = self.author_info_frame.winfo_width()
                height = self.author_info_frame.winfo_height()
                print(f"🔍 Author info position: x={x}, y={y}, w={width}, h={height}")
                
                # Check if labels exist
                if self.author_label1 and self.author_label1.winfo_exists():
                    print(f"🔍 Label1 text: '{self.author_label1.cget('text')}'")
                if self.author_label2 and self.author_label2.winfo_exists():
                    print(f"🔍 Label2 text: '{self.author_label2.cget('text')}'")
        except Exception as e:
            print(f"Debug error: {e}")
    
    def create_settings_button(self):
        """Tạo nút settings chỉ có biểu tượng bánh răng"""
        print("🔧 Creating gear icon only...")
        
        # Create settings container with place positioning
        self.settings_container = tk.Frame(self.root, bg=COLORS['bg_primary'])
        
        # Position above author info
        self.settings_container.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-45)

        # Create settings label (chỉ có biểu tượng, không có nền)
        self.settings_button = tk.Label(
            self.settings_container,
            text="⚙",
            font=('Segoe UI', 16, 'bold') if sys.platform.startswith('win') else ('Arial', 16, 'bold'),
            bg=COLORS['bg_primary'],   # Cùng màu nền với cửa sổ (trong suốt)
            fg='#888888',              # Màu xám nhạt cho bánh răng
            cursor="hand2"             # Con trỏ tay khi hover
        )
        self.settings_button.pack()
        
        # Bind click và hover events
        self.settings_button.bind("<Button-1>", lambda e: self.mo_file_mapping())
        self.settings_button.bind("<Enter>", self.on_settings_hover_enter)
        self.settings_button.bind("<Leave>", self.on_settings_hover_leave)
        
        # Bring to front
        self.settings_container.lift()
        print("✅ Gear icon created without background")
    
    def on_settings_hover_enter(self, event):
        """Hiệu ứng khi hover vào bánh răng"""
        self.settings_button.config(
            fg='#444444'       # Bánh răng đậm hơn khi hover
        )
    
    def on_settings_hover_leave(self, event):
        """Hiệu ứng khi rời khỏi bánh răng"""
        self.settings_button.config(
            fg='#888888'       # Trở về xám nhạt
        )
    
    def mo_file_mapping(self):
        """Mở file mapping.xlsx để chỉnh sửa"""
        mapping_path = get_mapping_file_path()
        
        if not open_file_with_system(mapping_path):
            # Fallback: show file location
            messagebox.showinfo(
                "Thông báo", 
                f"Không thể mở file mapping.xlsx tự động.\n\n"
                f"Vui lòng mở file tại:\n{mapping_path}\n\n"
                f"Sau khi chỉnh sửa xong, hãy lưu file.\n"
                f"Phần mềm sẽ tự động restart (không có console)."
            )
            return
        
        messagebox.showinfo(
            "Thông báo", 
            "Đã mở file mapping.xlsx.\n\n"
            "📝 LƯU Ý: Cấu trúc mới của Sheet2:\n"
            "A: apcu (ấp cũ)\n"
            "B: xacu (xã cũ)\n" 
            "C: huyencu (huyện cũ)\n"
            "D: tinhcu (tỉnh cũ)\n"
            "E: apmoi (ấp mới)\n"
            "F: xamoi (xã mới)\n"
            "G: tinhmoi (tỉnh mới)\n\n"
            "🔄 AUTO-RESTART:\n"
            "Sau khi lưu file và đóng Excel,\n"
            "phần mềm sẽ tự động khởi động lại\n"
            "(KHÔNG hiển thị console)."
        )
    
    def reset_ui(self):
        """Reset UI về trạng thái ban đầu"""
        self.main_window.processing = False
        self.main_window.paused = False
        self.main_window.stop_flag = False
        self.main_window.done_rows = 0
        self.main_window.total_paused_time = 0
        self.main_window.pause_start_time = 0
        
        # Reset multi-sheet variables
        self.main_window.current_sheet_index = 0
        self.main_window.total_sheets = 1
        self.main_window.sheet_results = {}
        
        self.label.config(text="Sẵn sàng xử lý danh sách bệnh nhân")
        self.progress['value'] = 0
        self.time_label.config(text="00:00 / 00:00 - thời gian đã xử lý / thời gian còn lại")
        
        self.control_frame.pack_forget()
        self.log_frame.pack_forget()
        
        # Just resize window without animation
        self.root.geometry(DEFAULT_GEOMETRY)
        
        # Show main button
        self.main_button_frame.pack(pady=20)
        
        # Ensure author info and settings are visible with place()
        if hasattr(self, 'author_info_frame') and self.author_info_frame:
            self.author_info_frame.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-15)
            self.author_info_frame.lift()
        
        if hasattr(self, 'settings_container') and self.settings_container:
            self.settings_container.place(relx=1.0, rely=1.0, anchor='se', x=-15, y=-45)
            self.settings_container.lift()
        
        self.pause_button.config(
            text="Tạm Dừng", 
            bg=COLORS['danger'],
            activebackground="#b71c1c"
        )
        
        print("✅ UI reset completed - author info should be visible")
    
    def update_sheet_log(self, message):
        """Cập nhật log với message cho sheet"""
        def update():
            try:
                self.log_box.config(state='normal')
                self.log_box.insert(tk.END, f"{message}\n")
                self.log_box.see(tk.END)
                self.log_box.config(state='disabled')
            except Exception as e:
                print(f"Sheet log update error: {e}")
        
        self.root.after(0, update)
    
    def update_log_with_sheet(self, chunk_index, chunk_total, current, total, sheet_index):
        """Cập nhật log xử lý với thông tin sheet"""
        def update():
            try:
                self.log_box.config(state='normal')
                lines = self.log_box.get("1.0", tk.END).strip().split("\n")
                
                # Ensure we have enough lines for all chunks across all sheets
                total_lines_needed = chunk_total * self.main_window.total_sheets
                while len(lines) < total_lines_needed:
                    lines.append("")

                # Calculate line index for this chunk in this sheet
                line_index = sheet_index * chunk_total + chunk_index
                
                if current >= total:
                    lines[line_index] = f"✅ Sheet {sheet_index+1} - Cụm {chunk_index+1}/{chunk_total} hoàn tất"
                else:
                    lines[line_index] = f"🔄 Sheet {sheet_index+1} - Cụm {chunk_index+1}/{chunk_total}: {current}/{total}"

                self.log_box.delete("1.0", tk.END)
                self.log_box.insert(tk.END, "\n".join(lines) + "\n")
                self.log_box.see(tk.END)
                self.log_box.config(state='disabled')
            except Exception as e:
                print(f"Log update error: {e}")

        self.root.after(0, update)
