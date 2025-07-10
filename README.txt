================================================================================
PIHCM – PHẦN MỀM CHUẨN HÓA ĐỊA CHỈ BỆNH NHÂN
================================================================================

📌 Phiên bản: 1.0 – Windows Edition (External Mapping)
📎 Đơn vị sở hữu: Viện Pasteur TP. HCM – Khoa KSPNBT
📞 Liên hệ kỹ thuật: Đinh Văn Ngôn – 08888 31135

--------------------------------------------------------------------------------
🎯 CHỨC NĂNG CHÍNH
--------------------------------------------------------------------------------
• Chuẩn hóa địa chỉ bệnh nhân theo đơn vị hành chính mới (sau 1/7/2025)
• Hỗ trợ định dạng Excel: .xlsx và .xls
• Tự động xử lý nhiều sheet trong cùng một file
• Matching thông minh bằng fuzzy logic
• Nhận diện và hỗ trợ ấp/thôn/khu phố (Sheet2)
• Tự động khởi động lại khi cập nhật mapping
• Sử dụng file mapping.xlsx ngoài, dễ dàng cập nhật theo từng địa phương

--------------------------------------------------------------------------------
🚀 CÁCH SỬ DỤNG PHẦN MỀM
--------------------------------------------------------------------------------
1. MỞ PHẦN MỀM
   • Nhấp đúp chuột vào file `.exe`
   • Phần mềm sẽ tự tìm file `mapping.xlsx` trong cùng thư mục

2. CHỌN FILE DỮ LIỆU CẦN CHUẨN HÓA
   • Click nút “Chọn file” hoặc kéo thả file Excel vào cửa sổ
   • File phải chứa 3 cột bắt buộc: Xã, Huyện, Tỉnh

--------------------------------------------------------------------------------
🛠️ CHỈNH SỬA FILE MAPPING
--------------------------------------------------------------------------------
🔧 CÁCH 1: CHỈNH SỬA TRỰC TIẾP TỪ GIAO DIỆN PHẦN MỀM
   1. Click nút ⚙ ở góc dưới bên phải
   2. File `mapping.xlsx` sẽ mở bằng Excel
   3. Chỉnh sửa và lưu file, sau đó đóng Excel
   4. Phần mềm sẽ tự động khởi động lại

🔧 CÁCH 2: THAY FILE MAPPING MỚI
   1. Tải file mapping.xlsx mới từ tác giả
   2. Backup file mapping.xlsx cũ (nếu cần)
   3. Thay thế file mapping.xlsx trong thư mục chứa .exe
   4. Khởi động lại chương trình

--------------------------------------------------------------------------------
📂 CẤU TRÚC FILE MAPPING.XLSX
--------------------------------------------------------------------------------
📄 SHEET1 – MAPPING ĐƠN VỊ HÀNH CHÍNH (Xã → Xã):
   - A: xacu     – Tên xã cũ
   - B: huyencu  – Tên huyện cũ
   - C: tinhcu   – Tên tỉnh cũ
   - D: xamoi    – Tên xã mới
   - E: tinhmoi  – Tên tỉnh mới

📄 SHEET2 – MAPPING ẤP/THÔN/KP:
   - A: apcu     – Tên ấp/thôn/khu phố cũ
   - B: xacu     – Tên xã cũ
   - C: huyencu  – Tên huyện cũ
   - D: tinhcu   – Tên tỉnh cũ
   - E: apmoi    – Tên ấp/thôn mới
   - F: xamoi    – Tên xã mới
   - G: tinhmoi  – Tên tỉnh mới

--------------------------------------------------------------------------------
💻 YÊU CẦU HỆ THỐNG
--------------------------------------------------------------------------------
• Windows 8, 8.1, 10, 11 (32-bit hoặc 64-bit)
• RAM: Tối thiểu 2GB (khuyến nghị 4GB trở lên)
• Dung lượng trống: ít nhất 100MB
• Cần cài đặt Microsoft Excel để chỉnh sửa mapping.xlsx

--------------------------------------------------------------------------------
📌 LƯU Ý QUẢN LÝ MAPPING
--------------------------------------------------------------------------------
• File `mapping.xlsx` PHẢI đặt cùng thư mục với file `.exe`
• Có thể sao lưu nhiều phiên bản mapping khác nhau
• Khi cập nhật mapping, chỉ cần thay file → chương trình sẽ tự nhận diện
• Sẽ báo lỗi nếu không tìm thấy hoặc sai định dạng file mapping.xlsx

--------------------------------------------------------------------------------
📞 HỖ TRỢ KỸ THUẬT
--------------------------------------------------------------------------------
Liên hệ: Đinh Văn Ngôn  
Đơn vị: Viện Pasteur TP. HCM – Khoa KSPNBT  
Điện thoại: 08888 31135  

--------------------------------------------------------------------------------
⚖️ BẢN QUYỀN
--------------------------------------------------------------------------------
Phần mềm thuộc quyền sở hữu của Viện Pasteur TP. HCM  
Chỉ sử dụng nội bộ, không được phát hành công khai hoặc thương mại hóa  

================================================================================
Cảm ơn bạn đã sử dụng phần mềm PIHCM!
================================================================================
