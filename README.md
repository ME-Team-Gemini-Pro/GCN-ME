# 🇻🇳 GCN-ME

> *"Help the Vietnamese state (in matters related to land) to operate more quickly and accurately."*  
> **Made by TonygamingTL**

![Python Version](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Author](https://img.shields.io/badge/author-TonygamingTL-orange)

**GCN-ME** (Giấy Chứng Nhận - Management & Extraction) là công cụ tự động hóa việc đọc, bóc tách và chuẩn hóa dữ liệu từ các tệp PDF hồ sơ địa chính, Giấy chứng nhận quyền sử dụng đất (Sổ đỏ/Sổ hồng) sang bảng biểu Excel (`.xlsx`), hỗ trợ cơ quan quản lý và cán bộ địa chính tại Việt Nam xử lý nghiệp vụ nhanh chóng, chính xác và giảm thiểu sai sót thủ công.

---

## 🎯 Mục tiêu dự án (Mission)

- **Tăng tốc độ xử lý hồ sơ:** Tự động hóa quá trình nhập liệu hàng loạt từ tệp PDF sang bảng tính Excel chỉ trong vài giây.
- **Độ chính xác cao:** Chuẩn hóa các trường thông tin theo đúng quy chuẩn dữ liệu địa chính Việt Nam, hạn chế tối đa nhầm lẫn khi gõ tay.
- **Tối ưu hóa quy trình:** Hỗ trợ tải trực tiếp hồ sơ từ Google Drive và xuất báo cáo tổng hợp sẵn sàng phục vụ công tác thống kê, tra cứu và lưu trữ.

---

## ✨ Tính năng chính (Key Features)

- **Trích xuất dữ liệu Giấy chứng nhận (GCN) tự động:**
  - Thông tin chủ sử dụng: Họ tên, năm sinh, số CCCD/Định danh cá nhân, địa chỉ.
  - Thông tin thửa đất: Số tờ bản đồ, số thửa, địa chỉ thửa đất, diện tích ($m^2$).
  - Thông tin pháp lý: Mục đích sử dụng đất (ONT, ODT, LUC, CLN...), thời hạn sử dụng, nguồn gốc sử dụng và số phát hành GCN.
- **Hỗ trợ đa định dạng PDF:**
  - Đọc trực tiếp PDF văn bản điện tử (Digital PDF).
  - Nhận diện ký tự quang học (OCR) tiếng Việt cho các hồ sơ quét ảnh (Scanned PDF).
- **Tích hợp lưu trữ đám mây:** Hỗ trợ kéo dữ liệu PDF hàng loạt từ Google Drive và tổng hợp thành 1 file Excel duy nhất.

---

## 🛠️ Cài đặt (Installation)

### 1. Yêu cầu hệ thống
- **Python:** 3.10 trở lên
- **Tesseract OCR** (kèm gói ngôn ngữ `vie` nếu xử lý hồ sơ scan ảnh)

### 2. Cài đặt mã nguồn ( CMD )
```bash
git clone [https://github.com/TonygamingTL/GCN-ME.git](https://github.com/TonygamingTL/GCN-ME.git)
cd GCN-ME
py -m pip install -r requirements.txt
```

Các thư viện chính trong `requirements.txt`:
```txt
pymupdf>=1.23.0
pdfplumber>=0.10.0
pandas>=2.1.0
openpyxl>=3.1.0
pytesseract>=0.3.10
gdown>=5.1.0
```

---

## 🚀 Hướng dẫn sử dụng (Usage)

### 1. Trích xuất từ thư mục PDF trên máy
Bỏ các file PDF Giấy chứng nhận vào thư mục `input_pdfs/` và chạy lệnh:

```bash
python main.py --input input_pdfs/ --output output/tong_hop_gcn.xlsx
```

### 2. Tải và trích xuất trực tiếp từ Google Drive
```bash
python main.py --gdrive "[https://drive.google.com/drive/folders/YOUR_FOLDER_ID](https://drive.google.com/drive/folders/YOUR_FOLDER_ID)" --output output/tong_hop_gcn.xlsx
```

---

## 📊 Cấu trúc bảng Excel đầu ra (Output Format)

Dữ liệu sau khi qua **GCN-ME** sẽ được xuất ra file `.xlsx` gồm các trường thông tin chuẩn:

1. **Mã hồ sơ / Tên file PDF** (`Ma_Ho_So`)
2. **Số phát hành GCN / Số vào sổ** (`So_Phat_Hanh`)
3. **Họ tên người sử dụng đất** (`Chu_Su_Dung`)
4. **Số CCCD / CMND** (`So_CCCD`)
5. **Thửa đất số** (`So_Thua`)
6. **Tờ bản đồ số** (`So_To_Ban_Do`)
7. **Địa chỉ thửa đất** (`Dia_Chi_Thua_Dat`)
8. **Diện tích ($m^2$)** (`Dien_Tich`)
9. **Mục đích sử dụng đất** (`Muc_Dich_Su_Dung`)
10. **Thời hạn sử dụng** (`Thoi_Han_Su_Dung`)
11. **Nguồn gốc sử dụng** (`Nguon_Goc_Su_Dung`)

---

## 🔒 Bảo mật thông tin (Security Notice)

Hồ sơ đất đai là tài liệu chứa thông tin công dân và tài sản quan trọng:
- Tuyệt đối không tải (push) các tệp PDF hồ sơ thật hoặc tệp Excel kết quả lên GitHub ở chế độ Public.
- Kiểm tra kỹ file `.gitignore` đã chặn các thư mục `input_pdfs/`, `output/`, `*.pdf`, và `*.xlsx` trước khi commit.

---

## 👤 Tác giả (Author)

- **Developed by:** **TonygamingTL**
- **Project:** GCN-ME — *Helping the Vietnamese state (in matters related to land) to operate more quickly and accurately.*

---

## 📄 Giấy phép (License)

Dự án được phân phối dưới giấy phép **MIT License**.
