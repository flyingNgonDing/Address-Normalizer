"""
File Watcher Module - Auto-restart application when mapping.xlsx is modified
Monitors mapping.xlsx file changes and triggers application restart
FIXED: No console window on restart
"""
import os
import sys
import time
import threading
import subprocess
import shutil
from pathlib import Path
from tkinter import messagebox

# Try to import watchdog with proper fallback
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
    print("✅ watchdog library loaded successfully")
except ImportError:
    WATCHDOG_AVAILABLE = False
    print("⚠️ watchdog not available. File watching disabled.")
    
    # Create dummy classes for fallback
    class FileSystemEventHandler:
        """Dummy FileSystemEventHandler when watchdog not available"""
        def __init__(self):
            pass
        
        def on_modified(self, event):
            pass
        
        def on_moved(self, event):
            pass
    
    class Observer:
        """Dummy Observer when watchdog not available"""
        def __init__(self):
            pass
        
        def schedule(self, handler, path, recursive=False):
            pass
        
        def start(self):
            pass
        
        def stop(self):
            pass
        
        def join(self, timeout=None):
            pass

from utils.helpers import get_mapping_file_path


class MappingFileHandler(FileSystemEventHandler):
    """Handles mapping.xlsx file modification events"""
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.mapping_path = get_mapping_file_path()
        self.mapping_filename = os.path.basename(self.mapping_path)
        self.last_modified = 0
        self.restart_pending = False
        self.excel_process_detected = False
        
        # Debounce settings to avoid multiple restarts
        self.debounce_time = 2.0  # Wait 2 seconds after last modification
        self.restart_timer = None
        
        if WATCHDOG_AVAILABLE:
            print(f"📁 Watching file: {self.mapping_path}")
        else:
            print(f"📁 File watching disabled: {self.mapping_path}")
    
    def on_modified(self, event):
        """Handle file modification events"""
        if not WATCHDOG_AVAILABLE:
            return
            
        if event.is_directory:
            return
        
        # Check if it's our mapping file
        if os.path.basename(event.src_path) == self.mapping_filename:
            self.handle_mapping_file_change(event.src_path)
    
    def on_moved(self, event):
        """Handle file move/rename events (Excel save behavior)"""
        if not WATCHDOG_AVAILABLE:
            return
            
        if event.is_directory:
            return
        
        # Excel often saves by creating temp file and renaming
        if os.path.basename(event.dest_path) == self.mapping_filename:
            self.handle_mapping_file_change(event.dest_path)
    
    def handle_mapping_file_change(self, file_path):
        """Handle mapping file changes with debouncing"""
        if not WATCHDOG_AVAILABLE:
            return
            
        current_time = time.time()
        
        # Skip if file doesn't exist or is being processed
        if not os.path.exists(file_path):
            return
        
        # Get file modification time
        try:
            file_mtime = os.path.getmtime(file_path)
        except OSError:
            return
        
        # Skip if this is the same modification we already processed
        if file_mtime <= self.last_modified:
            return
        
        # Check if Excel is still using the file
        if self.is_excel_using_file(file_path):
            print("📊 Excel is still using the file, waiting...")
            # Schedule a check in 1 second
            if self.restart_timer:
                self.restart_timer.cancel()
            self.restart_timer = threading.Timer(1.0, lambda: self.handle_mapping_file_change(file_path))
            self.restart_timer.start()
            return
        
        self.last_modified = file_mtime
        
        print(f"📝 Mapping file modified: {file_path}")
        print(f"🕐 Modification time: {file_mtime}")
        
        # Cancel any existing restart timer
        if self.restart_timer:
            self.restart_timer.cancel()
        
        # Start new debounced restart timer
        self.restart_timer = threading.Timer(self.debounce_time, self.trigger_restart)
        self.restart_timer.start()
        print(f"⏰ Restart scheduled in {self.debounce_time} seconds...")
    
    def is_excel_using_file(self, file_path):
        """Check if Excel is currently using the file"""
        try:
            # Method 1: Try to open file in exclusive mode
            with open(file_path, 'r+b') as f:
                pass
            return False
        except (IOError, OSError, PermissionError):
            # File is likely being used by Excel
            return True
    
    def trigger_restart(self):
        """Trigger application restart"""
        if not WATCHDOG_AVAILABLE:
            return
            
        if self.restart_pending:
            return
        
        self.restart_pending = True
        
        print("🔄 Triggering application restart...")
        
        # Show notification to user
        self.main_window.root.after(0, self.show_restart_notification)
        
        # Wait a bit for the notification to show
        threading.Timer(1.5, self.perform_restart).start()
    
    def show_restart_notification(self):
        """Show restart notification to user"""
        try:
            # Update main label to show restart status
            if hasattr(self.main_window, 'components') and self.main_window.components.label:
                self.main_window.components.label.config(
                    text="🔄 Phát hiện thay đổi mapping.xlsx - Đang khởi động lại...",
                    fg="#1976d2"  # Blue color
                )
            
            # Update window if possible
            self.main_window.root.update()
            
        except Exception as e:
            print(f"Error showing restart notification: {e}")
    
    def perform_restart(self):
        """Perform the actual restart - FIXED: No console window"""
        try:
            print("🚀 Restarting application...")
            
            # Create backup of current mapping file
            self.create_mapping_backup()
            
            # Get current executable path
            if getattr(sys, 'frozen', False):
                # Running as compiled exe
                exe_path = sys.executable
                print(f"Using executable path: {exe_path}")
            else:
                # Running as script - use python + script path
                exe_path = [sys.executable, sys.argv[0]]
                print(f"Using script path: {exe_path}")
            
            # Start new instance - FIXED: No console window
            if sys.platform.startswith('win'):
                try:
                    # FIXED: Use CREATE_NO_WINDOW instead of CREATE_NEW_CONSOLE
                    CREATE_NO_WINDOW = 0x08000000
                    
                    if isinstance(exe_path, list):
                        # For Python script
                        subprocess.Popen(exe_path, 
                                       creationflags=CREATE_NO_WINDOW,
                                       cwd=os.getcwd())
                    else:
                        # For executable
                        subprocess.Popen([exe_path], 
                                       creationflags=CREATE_NO_WINDOW,
                                       cwd=os.getcwd())
                    print("✅ New instance started without console window")
                    
                except Exception as e1:
                    print(f"Method 1 (no window) failed: {e1}")
                    try:
                        # Method 2: Try with DETACHED_PROCESS (no console)
                        DETACHED_PROCESS = 0x00000008
                        
                        if isinstance(exe_path, list):
                            subprocess.Popen(exe_path, 
                                           creationflags=DETACHED_PROCESS,
                                           cwd=os.getcwd())
                        else:
                            subprocess.Popen([exe_path], 
                                           creationflags=DETACHED_PROCESS,
                                           cwd=os.getcwd())
                        print("✅ New instance started with DETACHED_PROCESS")
                        
                    except Exception as e2:
                        print(f"Method 2 (detached) failed: {e2}")
                        try:
                            # Method 3: Try without special flags (fallback)
                            if isinstance(exe_path, list):
                                subprocess.Popen(exe_path, cwd=os.getcwd())
                            else:
                                subprocess.Popen([exe_path], cwd=os.getcwd())
                            print("✅ New instance started without special flags")
                            
                        except Exception as e3:
                            print(f"All methods failed: {e3}")
                            # Show error and don't close current instance
                            self.show_restart_error(str(e3))
                            return
            else:
                # Other platforms
                if isinstance(exe_path, list):
                    subprocess.Popen(exe_path)
                else:
                    subprocess.Popen([exe_path])
                print("✅ New instance started (non-Windows)")
            
            # Close current instance after short delay
            self.main_window.root.after(1000, self.main_window.root.quit)
            
        except Exception as e:
            print(f"Error during restart: {e}")
            self.show_restart_error(str(e))
    
    def show_restart_error(self, error_message):
        """Show restart error to user - FIXED lambda scope issue"""
        def show_error():
            try:
                messagebox.showerror(
                    "Lỗi khởi động lại", 
                    f"Không thể khởi động lại tự động:\n{error_message}\n\n"
                    f"Vui lòng khởi động lại phần mềm thủ công để áp dụng thay đổi mapping."
                )
            except Exception as msg_error:
                print(f"Error showing messagebox: {msg_error}")
            
            # Reset restart pending flag
            self.restart_pending = False
        
        # Schedule on main thread
        self.main_window.root.after(0, show_error)
    
    def create_mapping_backup(self):
        """Create backup of mapping file before restart"""
        try:
            backup_path = self.mapping_path + ".backup"
            shutil.copy2(self.mapping_path, backup_path)
            print(f"💾 Backup created: {backup_path}")
        except Exception as e:
            print(f"Warning: Could not create backup: {e}")


class FileWatcher:
    """Main file watcher class"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.observer = None
        self.handler = None
        self.watching = False
        
        if not WATCHDOG_AVAILABLE:
            print("⚠️ File watching not available (watchdog library missing)")
            return
        
        self.mapping_path = get_mapping_file_path()
        self.watch_directory = os.path.dirname(self.mapping_path)
        
        print(f"📂 Watch directory: {self.watch_directory}")
    
    def start_watching(self):
        """Start watching the mapping file"""
        if not WATCHDOG_AVAILABLE:
            print("⚠️ Cannot start file watching - watchdog library not available")
            print("   Install watchdog: pip install watchdog")
            return False
        
        if self.watching:
            return True
        
        try:
            # Ensure watch directory exists
            if not os.path.exists(self.watch_directory):
                print(f"❌ Watch directory does not exist: {self.watch_directory}")
                return False
            
            # Create event handler
            self.handler = MappingFileHandler(self.main_window)
            
            # Create observer
            self.observer = Observer()
            self.observer.schedule(self.handler, self.watch_directory, recursive=False)
            
            # Start watching
            self.observer.start()
            self.watching = True
            
            print(f"👁️ Started watching: {self.watch_directory}")
            print(f"📄 Monitoring file: {os.path.basename(self.mapping_path)}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error starting file watcher: {e}")
            return False
    
    def stop_watching(self):
        """Stop watching the mapping file"""
        if not self.watching or not self.observer:
            return
        
        try:
            # Cancel any pending restart timers
            if self.handler and self.handler.restart_timer:
                self.handler.restart_timer.cancel()
            
            # Stop observer
            self.observer.stop()
            self.observer.join(timeout=2.0)  # Wait up to 2 seconds
            
            self.watching = False
            print("👁️ Stopped file watching")
            
        except Exception as e:
            print(f"Warning: Error stopping file watcher: {e}")
    
    def is_watching(self):
        """Check if currently watching"""
        return self.watching and WATCHDOG_AVAILABLE


# Global file watcher instance
_file_watcher = None


def get_file_watcher():
    """Get the global file watcher instance"""
    return _file_watcher


def start_file_watching(main_window):
    """Start file watching for the main window"""
    global _file_watcher
    
    if not WATCHDOG_AVAILABLE:
        print("⚠️ File watching disabled - watchdog library not available")
        print("   To enable auto-restart: pip install watchdog")
        return False
    
    if _file_watcher is None:
        _file_watcher = FileWatcher(main_window)
    
    success = _file_watcher.start_watching()
    
    if success:
        print("✅ File watching enabled - Auto-restart when mapping.xlsx is saved")
    else:
        print("⚠️ File watching disabled - Manual restart required for mapping changes")
    
    return success


def stop_file_watching():
    """Stop file watching"""
    global _file_watcher
    
    if _file_watcher:
        _file_watcher.stop_watching()


def is_file_watching_active():
    """Check if file watching is currently active"""
    global _file_watcher
    return _file_watcher is not None and _file_watcher.is_watching()