================================================================================
PIHCM - PHẦN MỀM CHUẨN HÓA ĐỊA CHỈ BỆNH NHÂN
================================================================================
Sản phẩm thuộc sở hữu của: Viện Pasteur TP HCM - Khoa KSPNBT
Tác giả: Đinh Văn Ngôn (08888 31135)
Phiên bản: 1.0 - Windows Edition (External Mapping)

🎯 TÍNH NĂNG CHÍNH:
• Chuẩn hóa địa chỉ bệnh nhân theo đơn vị hành chính mới
• Hỗ trợ Excel (.xlsx, .xls) và CSV
• Xử lý nhiều sheet cùng lúc
• Matching thông minh với fuzzy logic
• Hỗ trợ ấp/thôn/khu phố (Sheet2)
• Auto-restart khi thay đổi mapping
• External mapping.xlsx có thể cập nhật được

📁 CẤU TRÚC GÓI PHẦN MỀM:
├── PIHCM - Chuyen doi don vi hanh chinh.exe  ← Chương trình chính (có icon)
├── mapping.xlsx                              ← File cấu hình mapping (EXTERNAL)
└── README.txt                                ← File hướng dẫn này

🔧 ƯU ĐIỂM EXTERNAL MAPPING:
✅ File mapping.xlsx ở bên ngoài - không nhúng vào .exe
✅ Có thể thay thế mapping.xlsx bằng phiên bản mới
✅ Người dùng có thể tự chỉnh sửa mapping
✅ Dễ dàng cập nhật mà không cần build lại phần mềm
✅ Backup và restore mapping dễ dàng

🚀 CÁCH SỬ DỤNG:
1. CHẠY CHƯƠNG TRÌNH:
   • Double-click vào file .exe
   • Icon tùy chỉnh sẽ hiển thị tự động (embedded trong .exe)
   • Chương trình sẽ tự động tìm file mapping.xlsx cùng thư mục

2. CHỌN FILE DANH SÁCH:
   • Click nút chọn file HOẶC kéo thả file vào cửa sổ
   • Hỗ trợ: .xlsx, .xls, .csv
   • File phải có các cột: Xã, Huyện, Tỉnh

⚙️ CHỈNH SỬA MAPPING:
🔧 CÁCH 1 - TỪ TRONG CHƯƠNG TRÌNH:
   1. Click nút ⚙ (bánh răng) ở góc dưới phải
   2. Excel sẽ mở file mapping.xlsx (EXTERNAL FILE)
   3. Chỉnh sửa theo cấu trúc
   4. Save file và đóng Excel
   5. Chương trình TỰ ĐỘNG RESTART ngay lập tức

🔧 CÁCH 2 - THAY THẾ FILE MAPPING:
   1. Tải file mapping.xlsx mới từ tác giả
   2. Backup file mapping.xlsx hiện tại (nếu cần)
   3. Thay thế file mapping.xlsx cùng thư mục với .exe
   4. Khởi động lại chương trình

📋 CẤU TRÚC MAPPING.XLSX:
📊 SHEET1 - Mapping cơ bản (Xã → Xã):
   Cột A: xacu     (Tên xã cũ)
   Cột B: huyencu  (Tên huyện cũ)  
   Cột C: tinhcu   (Tén tỉnh cũ)
   Cột D: xamoi    (Tên xã mới sau sáp nhập)
   Cột E: tinhmoi  (Tên tỉnh mới sau sáp nhập)

📊 SHEET2 - Mapping ấp/thôn (Ấp → Ấp):
   Cột A: apcu     (Tên ấp/thôn/KP cũ)
   Cột B: xacu     (Tên xã cũ)
   Cột C: huyencu  (Tên huyện cũ)
   Cột D: tinhcu   (Tên tỉnh cũ)
   Cột E: apmoi    (Tên ấp/thôn/KP mới)
   Cột F: xamoi    (Tên xã mới)
   Cột G: tinhmoi  (Tỉnh mới)

💻 YÊU CẦU HỆ THỐNG:
• Windows 8/8.1/10/11 (32-bit hoặc 64-bit)
• RAM: Tối thiểu 2GB (khuyến nghị 4GB+)
• Dung lượng: 100MB trống
• Microsoft Excel (để chỉnh sửa mapping)

🔧 QUẢN LÝ MAPPING:
• File mapping.xlsx PHẢI cùng thư mục với .exe
• Có thể backup nhiều phiên bản mapping khác nhau
• Khi cập nhật mapping, chỉ cần thay file và restart
• Chương trình sẽ báo lỗi nếu không tìm thấy mapping.xlsx

📞 HỖ TRỢ KỸ THUẬT:
• Tác giả: Đinh Văn Ngôn
• Điện thoại: 08888 31135
• Đơn vị: Viện Pasteur TP HCM - Khoa KSPNBT

⚖️ BẢN QUYỀN:
Phần mềm thuộc sở hữu của Viện Pasteur TP HCM
Chỉ sử dụng nội bộ - Không phân phối lại

================================================================================
Cảm ơn bạn đã sử dụng phần mềm PIHCM!
================================================================================