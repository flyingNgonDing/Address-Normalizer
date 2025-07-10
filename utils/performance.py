"""
Utilities cho performance và threading
"""
import multiprocessing
from config import CHUNK_SIZE, MAX_WORKERS


def detect_mode():
    """
    Phát hiện chế độ xử lý tối ưu dựa trên số CPU cores
    
    Returns:
        str: "serial", "chunked", hoặc "parallel"
    """
    cores = multiprocessing.cpu_count()
    if cores <= 2:
        return "serial"
    elif cores <= 4:
        return "chunked"
    else:
        return "parallel"


def calculate_optimal_chunk_size(total_rows, num_workers=None):
    """
    Tính chunk size tối ưu dựa trên số lượng dữ liệu và workers
    
    Args:
        total_rows: Tổng số dòng dữ liệu
        num_workers: Số worker threads
        
    Returns:
        int: Chunk size tối ưu
    """
    if num_workers is None:
        num_workers = min(MAX_WORKERS, multiprocessing.cpu_count())
    
    # Đảm bảo mỗi worker có ít nhất 1 chunk
    if total_rows < num_workers:
        return max(1, total_rows // num_workers)
    
    # Tính chunk size để có ít nhất 2 chunks per worker
    optimal_size = total_rows // (num_workers * 2)
    
    # Giới hạn trong khoảng hợp lý
    return max(100, min(CHUNK_SIZE, optimal_size))


def get_worker_count():
    """
    Lấy số lượng worker threads tối ưu
    
    Returns:
        int: Số worker threads
    """
    return min(MAX_WORKERS, multiprocessing.cpu_count())


class ProcessingStats:
    """Class theo dõi thống kê xử lý"""
    
    def __init__(self):
        self.reset()
    
    def reset(self):
        """Reset tất cả thống kê"""
        self.total_rows = 0
        self.processed_rows = 0
        self.start_time = None
        self.end_time = None
        self.paused_time = 0
        self.errors = []
    
    def get_progress_percentage(self):
        """Lấy % tiến độ"""
        if self.total_rows == 0:
            return 0
        return (self.processed_rows / self.total_rows) * 100
    
    def get_processing_speed(self, current_time):
        """Tính tốc độ xử lý (rows/second)"""
        if not self.start_time or self.processed_rows == 0:
            return 0
        
        elapsed = current_time - self.start_time - self.paused_time
        if elapsed <= 0:
            return 0
        
        return self.processed_rows / elapsed
    
    def estimate_remaining_time(self, current_time):
        """Ước tính thời gian còn lại"""
        speed = self.get_processing_speed(current_time)
        if speed <= 0:
            return 0
        
        remaining_rows = self.total_rows - self.processed_rows
        return remaining_rows / speed