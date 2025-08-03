"""
Animation Handler
Manages all animations for the main window
"""
import time
import sys
from config import DEFAULT_GEOMETRY, EXPANDED_GEOMETRY


class AnimationHandler:
    """Handles all animation operations"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.root = main_window.root
        self.components = None  # Will be set after components are created
        
        # Animation state
        self.animation_running = False
        self.is_expanded = False
    
    def set_components(self, components):
        """Set reference to window components"""
        self.components = components
    
    def animate_window_resize(self):
        """Animate window resize for Windows"""
        try:
            current_geom = self.root.geometry()
            current_parts = current_geom.split('x')
            current_width = int(current_parts[0])
            current_height = int(current_parts[1].split('+')[0])
            
            target_parts = EXPANDED_GEOMETRY.split('x')
            target_width = int(target_parts[0])
            target_height = int(target_parts[1])
            
            steps = 8  # Fewer steps for smoother performance on Windows
            
            for i in range(steps + 1):
                progress = i / steps
                new_width = int(current_width + (target_width - current_width) * progress)
                new_height = int(current_height + (target_height - current_height) * progress)
                
                self.root.geometry(f"{new_width}x{new_height}")
                self.root.update()
                time.sleep(0.02)
                
        except Exception as e:
            print(f"Window resize animation error: {e}")
            self.root.geometry(EXPANDED_GEOMETRY)
    
    def animate_window_resize_back(self):
        """Animate window resize back to original size"""
        try:
            current_geom = self.root.geometry()
            current_parts = current_geom.split('x')
            current_width = int(current_parts[0])
            current_height = int(current_parts[1].split('+')[0])
            
            target_parts = DEFAULT_GEOMETRY.split('x')
            target_width = int(target_parts[0])
            target_height = int(target_parts[1])
            
            steps = 6
            
            for i in range(steps + 1):
                progress = i / steps
                new_width = int(current_width + (target_width - current_width) * progress)
                new_height = int(current_height + (target_height - current_height) * progress)
                
                self.root.geometry(f"{new_width}x{new_height}")
                self.root.update()
                time.sleep(0.02)
                
        except Exception as e:
            print(f"Window resize back animation error: {e}")
            self.root.geometry(DEFAULT_GEOMETRY)
    
    def animate_shape(self, target_coords, target_text, target_width):
        """Animate shape transformation - Windows optimized"""
        if self.animation_running or not self.main_window.use_animations:
            return
            
        self.animation_running = True
        
        # Get components reference
        components = self.main_window.components
        if not components or not hasattr(components, 'settings_canvas'):
            self.animation_running = False
            return
            
        start_coords = components.current_coords.copy()
        start_width = components.settings_canvas.winfo_width()
        
        try:
            for step in range(self.main_window.animation_steps + 1):
                if not self.animation_running:
                    break
                    
                progress = step / self.main_window.animation_steps
                eased_progress = self.main_window.animation_helper.ease_in_out_quad(progress)
                
                # Interpolate coordinates và width
                new_coords = self.main_window.animation_helper.interpolate_coords(
                    start_coords, target_coords, eased_progress
                )
                new_width = int(start_width + (target_width - start_width) * eased_progress)
                
                # Update canvas và shape
                components.settings_canvas.config(width=new_width)
                components.settings_canvas.delete("all")
                
                # Vẽ shape
                self._draw_animated_shape(new_coords, target_text)
                
                components.current_coords = new_coords.copy()
                
                # Update GUI
                try:
                    self.root.update()
                except:
                    break
                
                if step < self.main_window.animation_steps:
                    time.sleep(self.main_window.animation_delay / 1000.0)
                    
        except Exception as e:
            print(f"Animation error: {e}")
        finally:
            self.animation_running = False
    
    def _draw_animated_shape(self, coords, text):
        """Vẽ shape cho animation - Windows optimized"""
        try:
            components = self.main_window.components
            if not components or not hasattr(components, 'settings_canvas'):
                return
                
            width = abs(coords[2] - coords[0])
            height = abs(coords[3] - coords[1])
            
            if width > height * 1.8:  # Wide rectangle
                corner_radius = height / 2
                
                # Draw rounded rectangle using multiple elements
                # Left circle
                left_circle = [coords[0], coords[1], coords[0] + height, coords[3]]
                components.settings_canvas.create_oval(*left_circle, fill='#ffffff', outline='#bdc3c7', width=1)
                
                # Right circle
                right_circle = [coords[2] - height, coords[1], coords[2], coords[3]]
                components.settings_canvas.create_oval(*right_circle, fill='#ffffff', outline='#bdc3c7', width=1)
                
                # Center rectangle
                if width > height:
                    rect_coords = [coords[0] + corner_radius, coords[1], coords[2] - corner_radius, coords[3]]
                    components.settings_canvas.create_rectangle(*rect_coords, fill='#ffffff', outline='#ffffff')
                    
                    # Top and bottom borders
                    components.settings_canvas.create_line(
                        coords[0] + corner_radius, coords[1],
                        coords[2] - corner_radius, coords[1],
                        fill='#bdc3c7'
                    )
                    components.settings_canvas.create_line(
                        coords[0] + corner_radius, coords[3],
                        coords[2] - corner_radius, coords[3],
                        fill='#bdc3c7'
                    )
            else:
                # Circle/oval
                components.settings_canvas.create_oval(*coords, fill='#ffffff', outline='#bdc3c7', width=1)
            
            # Add text
            center_x = (coords[0] + coords[2]) / 2
            center_y = (coords[1] + coords[3]) / 2
            
            font_size = 8 if text != '⚙' else 12
            font_family = 'Segoe UI' if sys.platform.startswith('win') else 'Arial'
            
            components.text_id = components.settings_canvas.create_text(
                center_x, center_y,
                text=text,
                font=(font_family, font_size, 'bold'),
                fill='#7f8c8d'
            )
        except Exception as e:
            print(f"Error drawing shape: {e}")
    
    def on_settings_enter(self, event):
        """Expand to rectangle on hover"""
        components = self.main_window.components
        if not components:
            return
            
        if not self.is_expanded and not self.animation_running and self.main_window.use_animations:
            self.is_expanded = True
            self.animate_shape(components.expanded_coords, "Chỉnh sửa mapping", 175)

    def on_settings_leave(self, event):
        """Contract to circle when leaving"""
        if not self.main_window.use_animations:
            return
            
        components = self.main_window.components
        if not components:
            return
            
        try:
            x, y = self.root.winfo_pointerxy()
            widget_x = components.settings_canvas.winfo_rootx()
            widget_y = components.settings_canvas.winfo_rooty()
            widget_width = components.settings_canvas.winfo_width()
            widget_height = components.settings_canvas.winfo_height()
            
            if not (widget_x <= x <= widget_x + widget_width and 
                    widget_y <= y <= widget_y + widget_height):
                if self.is_expanded and not self.animation_running:
                    self.is_expanded = False
                    self.animate_shape(components.circle_coords, "⚙", 35)
        except:
            # Fallback: always contract
            if self.is_expanded and not self.animation_running:
                self.is_expanded = False
                self.animate_shape(components.circle_coords, "⚙", 35)