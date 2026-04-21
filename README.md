# 🛠 BÁO CÁO BẢO TRÌ BĂNG CHUYỀN TỰ ĐỘNG (BHS Maintenance Report Generator)

Hệ thống tự động hóa thao tác trích xuất, gom nhóm và kết xuất báo cáo lịch sử sửa chữa/bảo trì băng chuyền tại nhà ga. Nó nhặt dữ liệu thô (được trộn lẫn nhiều sheet và hay bị gộp ô phức tạp) và biến chúng thành 1 file báo cáo 4 Sheet cực kỳ chuyên nghiệp.

## 📂 1. Cấu Trúc Thư Mục
Hệ thống sử dụng các tên file chuẩn mực (tiếng Anh) như sau:
* **`source_mds_maintenance.xlsx`**: File nguồn NHẬT KÝ MDS (chứa sheet Sổ theo dõi CTBT).
* **`source_eq_history.xlsx`**: File nguồn LÝ LỊCH THIẾT BỊ (chứa sheet Sổ theo dõi CTBT).
* **`reference_conveyor_symbols.xlsx`**: File tài liệu tham khảo danh mục ký hiệu băng chuyền (Không bắt buộc chạy code).
* **`generate_report.py`**: Mã nguồn Python mang trí tuệ AI phân loại. Đóng vai trò TRÁI TIM của hệ thống.
* **`final_equipment_report.xlsx`** (Output sinh ra): Báo cáo gộp tuyệt đẹp phân loại sẵn vào 4 Sheets.

## ⚙️ 2. Hệ Thống "Smart Categorize" (Phân loại thông minh)
Hệ thống ngầm sử dụng **Regular Expressions (Regex)** để bắt dính mã băng chuyền từ nội dung nhập liệu tay và tự động vứt vào 4 khu vực:
* L, ME, SM1, SM2, SM3, SM4 ➡️ `BẢNG 1_BC ĐI`
* A, MCP A ➡️ `BẢNG 2_BC ĐẾN`
* Check-in, Bàn cân I7, Máy soi... ➡️ `BẢNG 3_CHECK IN`
* Sorter, Induction, Chute, SM5... ➡️ `BẢNG 4_SORTER`

## 💻 3. Hướng Dẫn Sử Dụng
**Bước 1: Cài đặt thư viện**
Chắc chắn bạn đã cài đặt Python 3. Sau đó mở Terminal và chạy lệnh cài đặt 2 plugin cần thiết để đọc/ghi file Excel:
```bash
pip install pandas openpyxl
```

**Bước 2: Cập nhật file nguồn**
Vào cuối ca trực, chỉ việc lưu 2 file nguồn (`source_eq_history.xlsx` và `source_mds_maintenance.xlsx`) có chứa dữ liệu mới nhất vào chung thư mục chứa file Code. ĐẾN ĐÂY BẠN XONG VIỆC, KHÔNG CẦN FOTMAT GÌ CẢ.

**Bước 3: Chạy Lệnh Tạo Báo Cáo**
Dùng Terminal (hoặc cmd/powershell), trỏ đường dẫn về thư mục này và gõ đúng 1 lệnh duy nhất:
```bash
python3 generate_report.py
```

**Bước 4: Nhận Báo Cáo**
Chỉ 1 giây sau, hệ thống sẽ in ra thông báo `THÀNH CÔNG`. 
Lúc này bạn sẽ thấy xuất hiện file **`final_equipment_report.xlsx`**. Mở lên, ctrl+p (In ra) và đi nộp cho Sếp!

---
*Created carefully by your AI Assistant ❤️.*
