"""
UI styling và themes cho application
WINDOWS VERSION - Optimized for Windows 8/10/11
"""
from tkinter import font, ttk
import sys
from config import COLORS, ANIMATION_SETTINGS


class StyleManager:
    """Class quản lý styles cho UI - Windows optimized"""
    
    def __init__(self):
        self.fonts = {}
        self.style = None
        self._fonts_initialized = False
        self.is_windows = sys.platform.startswith('win')
    
    def _setup_fonts(self):
        """Thiết lập các font chữ - Windows optimized"""
        if self._fonts_initialized:
            return
            
        try:
            if self.is_windows:
                # Windows specific fonts với size lớn hơn
                self.fonts = {
                    'custom': font.Font(family="Segoe UI", size=12),      # Tăng từ 10 → 12
                    'main_button': font.Font(family="Segoe UI", size=16, weight="bold"),  # Tăng từ 12 → 16
                    'small': font.Font(family="Segoe UI", size=10),      # Tăng từ 8 → 10
                    'console': font.Font(family="Consolas", size=10),    # Tăng từ 8 → 10
                    'header': font.Font(family="Segoe UI", size=14, weight="bold"),  # Tăng từ 11 → 14
                }
            else:
                # Fallback fonts for other systems với size lớn hơn
                self.fonts = {
                    'custom': font.Font(family="Arial", size=12),
                    'main_button': font.Font(family="Arial", size=16, weight="bold"),
                    'small': font.Font(family="Arial", size=10),
                    'console': font.Font(family="Courier", size=10),
                    'header': font.Font(family="Arial", size=14, weight="bold"),
                }
            self._fonts_initialized = True
        except Exception as e:
            print(f"Warning: Could not create custom fonts: {e}")
            # Fallback to default fonts
            self.fonts = {
                'custom': None,
                'main_button': None,
                'small': None,
                'console': None,
                'header': None,
            }
            self._fonts_initialized = True
    
    def setup_ttk_styles(self):
        """Thiết lập TTK styles - Windows optimized"""
        if not self._fonts_initialized:
            self._setup_fonts()
            
        self.style = ttk.Style()
        
        # Use Windows native theme if available
        if self.is_windows:
            try:
                available_themes = self.style.theme_names()
                if 'winnative' in available_themes:
                    self.style.theme_use('winnative')
                elif 'vista' in available_themes:
                    self.style.theme_use('vista')
                elif 'xpnative' in available_themes:
                    self.style.theme_use('xpnative')
                else:
                    self.style.theme_use('clam')
            except:
                self.style.theme_use('clam')
        else:
            self.style.theme_use('clam')
        
        # Configure button styles
        self._setup_button_styles()
        
        # Configure progress bar style
        self._setup_progressbar_style()
        
        # Configure other widgets
        self._setup_other_widgets()
    
    def _setup_button_styles(self):
        """Thiết lập styles cho buttons - Windows optimized với contrast cao"""
        # Windows specific button styling với font size lớn hơn
        button_font = ("Segoe UI", 12) if self.is_windows else ("Arial", 12)  # Tăng từ 10 → 12
        main_button_font = ("Segoe UI", 16, "bold") if self.is_windows else ("Arial", 16, "bold")  # Tăng từ 14 → 16
        
        # Red button style
        self.style.configure("Red.TButton", 
                           foreground="white", 
                           background=COLORS['danger'],
                           font=button_font, 
                           padding=(15, 10), 
                           relief="raised",
                           borderwidth=2,
                           focuscolor='none')
        self.style.map("Red.TButton", 
                     background=[("active", "#b71c1c"), ("pressed", "#c62828")],
                     relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Green button style
        self.style.configure("Green.TButton", 
                           foreground="white", 
                           background=COLORS['success'],
                           font=button_font, 
                           padding=(15, 10), 
                           relief="raised",
                           borderwidth=2,
                           focuscolor='none')
        self.style.map("Green.TButton", 
                     background=[("active", "#1b5e20"), ("pressed", "#2e7d32")],
                     relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Blue button style
        self.style.configure("Blue.TButton", 
                           foreground="white", 
                           background=COLORS['accent'],
                           font=button_font, 
                           padding=(15, 10), 
                           relief="raised",
                           borderwidth=2,
                           focuscolor='none')
        self.style.map("Blue.TButton", 
                     background=[("active", "#0d47a1"), ("pressed", "#1565c0")],
                     relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Main button style - XANH DƯƠNG ĐẬM
        self.style.configure("Main.TButton", 
                           foreground="white", 
                           background=COLORS['main_button'],  # Sử dụng main_button color
                           font=main_button_font, 
                           padding=(30, 15),      # Tăng padding
                           relief="raised",
                           borderwidth=3,
                           focuscolor='none')
        self.style.map("Main.TButton", 
                     background=[("active", "#1565c0"), ("pressed", "#0a3d91")],  # Darker shades
                     relief=[("pressed", "sunken"), ("!pressed", "raised")])
    
    def _setup_progressbar_style(self):
        """Thiết lập style cho progress bar - Windows optimized với contrast cao"""
        self.style.configure("Modern.Horizontal.TProgressbar",
                           background=COLORS['accent'],
                           troughcolor='#e0e0e0',
                           borderwidth=2,
                           lightcolor=COLORS['accent'],
                           darkcolor=COLORS['accent'],
                           relief='solid')
        
        # Windows specific progress bar styling
        if self.is_windows:
            self.style.configure("Modern.Horizontal.TProgressbar",
                               thickness=28)  # Tăng thickness để phù hợp với font lớn hơn
    
    def _setup_other_widgets(self):
        """Setup styles for other widgets"""
        # Frame styles
        self.style.configure("Card.TFrame",
                           background=COLORS['bg_primary'],
                           relief='flat',
                           borderwidth=1)
        
        # Label styles
        self.style.configure("Header.TLabel",
                           background=COLORS['bg_primary'],
                           foreground=COLORS['text_primary'],
                           font=self.get_font('header'))
        
        self.style.configure("Secondary.TLabel",
                           background=COLORS['bg_primary'],
                           foreground=COLORS['text_secondary'],
                           font=self.get_font('small'))
    
    def get_font(self, font_name):
        """Lấy font theo tên"""
        if not self._fonts_initialized:
            self._setup_fonts()
        return self.fonts.get(font_name, self.fonts.get('custom', None))
    
    def apply_theme(self, widget, theme_name):
        """Áp dụng theme cho widget"""
        themes = {
            'primary': {
                'bg': COLORS['bg_primary'],
                'fg': COLORS['text_primary']
            },
            'secondary': {
                'bg': COLORS['bg_primary'],
                'fg': COLORS['text_secondary']
            },
            'drag_active': {
                'bg': COLORS['bg_drag'],
                'fg': '#1976d2'
            },
            'success': {
                'bg': COLORS['success'],
                'fg': 'white'
            },
            'danger': {
                'bg': COLORS['danger'],
                'fg': 'white'
            }
        }
        
        if theme_name in themes:
            theme = themes[theme_name]
            try:
                widget.configure(**theme)
            except Exception as e:
                print(f"Warning: Could not apply theme {theme_name}: {e}")
    
    def configure_text_widget(self, text_widget):
        """Configure text widget for Windows"""
        try:
            if self.is_windows:
                text_widget.configure(
                    font=self.get_font('console'),
                    bg='#ffffff',
                    fg='#2c3e50',
                    relief='solid',
                    borderwidth=1,
                    highlightthickness=0,
                    selectbackground='#3498db',
                    selectforeground='white',
                    wrap='word'
                )
            else:
                text_widget.configure(
                    font=self.get_font('console'),
                    bg='#ffffff',
                    fg='#212529',
                    relief='solid',
                    borderwidth=1,
                    highlightthickness=0,
                    selectbackground='#e3f2fd'
                )
        except Exception as e:
            print(f"Warning: Could not configure text widget: {e}")


class AnimationHelper:
    """Helper class cho animations - Windows optimized"""
    
    @staticmethod
    def get_animation_settings():
        """Lấy settings cho animation"""
        # Reduce animation complexity on Windows for better performance
        settings = ANIMATION_SETTINGS.copy()
        if sys.platform.startswith('win'):
            settings['steps'] = max(5, settings['steps'] // 2)  # Fewer steps
            settings['delay'] = min(50, settings['delay'] + 10)  # Longer delay
        return settings
    
    @staticmethod
    def ease_out_cubic(progress):
        """Easing function cho smooth animation"""
        return 1 - (1 - progress) ** 3
    
    @staticmethod
    def ease_in_out_quad(progress):
        """Alternative easing function for Windows"""
        if progress < 0.5:
            return 2 * progress * progress
        else:
            return -1 + (4 - 2 * progress) * progress
    
    @staticmethod
    def interpolate_coords(start_coords, end_coords, progress):
        """Interpolate coordinates cho animation"""
        result = []
        for i in range(len(start_coords)):
            start_val = start_coords[i]
            end_val = end_coords[i]
            new_val = start_val + (end_val - start_val) * progress
            result.append(new_val)
        return result
    
    @staticmethod
    def should_use_animation():
        """Check if animations should be used on this system"""
        # Disable complex animations on older Windows versions
        if sys.platform.startswith('win'):
            try:
                import platform
                version = platform.release()
                return version in ['10', '11']  # Only enable on Win 10/11
            except:
                return False
        return True


# Global style manager instance
_style_manager = None


def get_style_manager():
    """Lấy style manager instance - tạo khi cần"""
    global _style_manager
    if _style_manager is None:
        _style_manager = StyleManager()
    return _style_manager


def setup_ui_styles():
    """Setup tất cả UI styles - Windows optimized"""
    style_manager = get_style_manager()
    style_manager.setup_ttk_styles()
    return style_manager


def apply_windows_theme(root):
    """Apply Windows-specific theme settings"""
    try:
        if sys.platform.startswith('win'):
            # Set Windows-specific options
            root.option_add('*TCombobox*Listbox.selectBackground', COLORS['accent'])
            root.option_add('*TCombobox*Listbox.font', 'Segoe UI 9')
            
            # Configure for high DPI displays
            try:
                root.tk.call('tk', 'scaling', 1.0)
            except:
                pass
                
    except Exception as e:
        print(f"Warning: Could not apply Windows theme: {e}")