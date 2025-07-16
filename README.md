# 🏥 PIHCM – Phần mềm chuẩn hóa địa chỉ bệnh nhân

**Phiên bản:** 1.0 – *Windows Edition*  
**Đơn vị:** Viện Pasteur TP. HCM – Khoa KSPNBT  
**Liên hệ kỹ thuật:** Đinh Văn Ngôn – 📞 08888 31135

---

## 🔽 Tải xuống phần mềm

🆕 **Người dùng lần đầu nên tải gói đầy đủ (.zip):**

👉 [📦 Chuan.Hoa.Don.Vi.Hanh.Chinh.v1.0.zip](https://github.com/flyingNgonDing/Address-Normalizer/releases/download/v1.0/Chuan.Hoa.Don.Vi.Hanh.Chinh.v1.0.zip)  
*(Gồm phần mềm `.exe` + file dữ liệu `mapping.xlsx`)*

🔸 Hoặc tải riêng từng thành phần:

- [💻 Phần mềm (.exe)](https://github.com/flyingNgonDing/Address-Normalizer/releases/download/v1.0/PIHCM.-.Chuyen.doi.don.vi.hanh.chinh.exe)  
- [📊 File mapping.xlsx](https://github.com/flyingNgonDing/Address-Normalizer/releases/download/v1.0/mapping.xlsx)

---

## 📖 Mục lục

- [🎯 Chức năng chính](#-chức-năng-chính)
- [🚀 Cách sử dụng](#-cách-sử-dụng)
- [⚙️ Chỉnh sửa mapping](#️-chỉnh-sửa-mapping)
- [📂 Cấu trúc mapping.xlsx](#-cấu-trúc-mappingxlsx)
- [💻 Yêu cầu hệ thống](#-yêu-cầu-hệ-thống)
- [📞 Hỗ trợ kỹ thuật](#-hỗ-trợ-kỹ-thuật)
- [⚖️ Bản quyền](#-bản-quyền)

---

## 🎯 Chức năng chính

- ✅ Chuẩn hóa địa chỉ xã/huyện/tỉnh theo đơn vị hành chính mới (sau ngày 01/07/2025)
- 📂 Hỗ trợ định dạng Excel `.xlsx`, `.xls`
- 📑 Xử lý nhiều sheet Excel trong cùng một lần
- 🤖 Matching thông minh với thuật toán fuzzy logic
- 🧩 Hỗ trợ địa danh cấp ấp/thôn/khu phố thông qua Sheet2
- 🔄 Tự động khởi động lại khi file mapping.xlsx được cập nhật
- 📁 Dễ dàng thay thế hoặc điều chỉnh file mapping.xlsx ngoài chương trình

---

## 🚀 Cách sử dụng

1. **Chạy chương trình:**
   - Nhấp đúp vào file `.exe`
   - Phần mềm sẽ tự tìm file `mapping.xlsx` trong cùng thư mục

2. **Chọn file dữ liệu Excel cần xử lý:**
   - Click nút chọn file hoặc kéo thả vào giao diện
   - Yêu cầu file phải có 3 cột: `Xã`, `Huyện`, `Tỉnh`

3. **Kết quả:**
   - Phần mềm xuất ra 3 file:
     - `ketqua_chuan.xlsx`: danh sách đã chuẩn hóa
     - `xacauveo_...xlsx`: xã cấu véo
     - `khongmatch_...xlsx`: không thể chuẩn hóa

---

## ⚙️ Chỉnh sửa mapping

### Cách 1 – Từ giao diện phần mềm:
- Click nút ⚙️ (bánh răng)
- Excel sẽ mở `mapping.xlsx`
- Chỉnh sửa và **Save**
- Phần mềm sẽ **tự động khởi động lại**

### Cách 2 – Thay file mới:
- Tải `mapping.xlsx` mới từ tác giả
- Ghi đè file cũ trong thư mục chứa `.exe`
- Mở lại phần mềm

---

## 📂 Cấu trúc mapping.xlsx

### 📄 Sheet1 – Đơn vị hành chính (Xã → Xã)
| Cột | Nội dung         |
|------|------------------|
| A    | `xacu` – Tên xã cũ |
| B    | `huyencu` – Tên huyện cũ |
| C    | `tinhcu` – Tên tỉnh cũ |
| D    | `xamoi` – Tên xã mới |
| E    | `tinhmoi` – Tên tỉnh mới |

### 📄 Sheet2 – Thôn/ấp/khu phố (Ấp → Ấp)
| Cột | Nội dung          |
|------|-------------------|
| A    | `apcu` – Ấp/thôn/KP cũ |
| B    | `xacu` – Tên xã cũ     |
| C    | `huyencu` – Huyện cũ   |
| D    | `tinhcu` – Tỉnh cũ     |
| E    | `apmoi` – Ấp/thôn mới  |
| F    | `xamoi` – Xã mới       |
| G    | `tinhmoi` – Tỉnh mới   |

---

## 💻 Yêu cầu hệ thống

- Hệ điều hành: Windows 8/8.1/10/11 (32-bit hoặc 64-bit)
- RAM: Tối thiểu 2GB (khuyến nghị 4GB trở lên)
- Bộ nhớ trống: 100MB
- **Microsoft Excel**: bắt buộc để mở chỉnh sửa mapping

---

## 📞 Hỗ trợ kỹ thuật

**Liên hệ:** Đinh Văn Ngôn  
**Đơn vị:** Viện Pasteur TP. HCM – Khoa KSPNBT  
**Điện thoại:** 08888 31135

---

## ⚖️ Bản quyền

Phần mềm thuộc quyền sở hữu của **Viện Pasteur TP. HCM**  
🔒 Chỉ sử dụng nội bộ – **Không phân phối lại, không thương mại hóa**

---

> 🙏 **Cảm ơn bạn đã sử dụng phần mềm PIHCM!**
> ![CodeRabbit Pull Request Reviews](https://img.shields.io/coderabbit/prs/github/flyingNgonDing/Address-Normalizer?utm_source=oss&utm_medium=github&utm_campaign=flyingNgonDing%2FAddress-Normalizer&labelColor=171717&color=FF570A&link=https%3A%2F%2Fcoderabbit.ai&label=CodeRabbit+Reviews)
