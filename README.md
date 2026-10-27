# HCMK26CNTT2_IT108_Session10_Ex04
Bài tập số 4
# 🚀 Bộ Hồ Sơ Thiết Kế Hệ Thống Quản Lý Bệnh Án Điện Tử Clinic-Cloud (Session 10 Ex04)

Bài giải chuẩn hóa bộ tài liệu đặc tả SRS, quản lý dữ liệu bệnh nhân và báo cáo dịch bệnh cho Phòng khám Startup Clinic-Cloud.

---

# Phần 1. Phân Tích Yêu Cầu Chức Năng và Phi Chức Năng (Functional vs Non-Functional Requirements)

## 1. Phân biệt Yêu cầu Chức năng và Phi chức năng

- **Yêu cầu Chức năng (Functional Requirements):** Là những tính năng trực tiếp phục vụ cho nghiệp vụ khám chữa bệnh hằng ngày của y bác sĩ và nhân viên y tế (Hệ thống làm được gì?).
- **Yêu cầu Phi chức năng (Non-Functional Requirements):** Là các tiêu chuẩn chất lượng kỹ thuật, quy định bảo mật, hiệu năng và tính tuân thủ pháp lý của hệ thống (Hệ thống vận hành an toàn/nhanh ra sao?).

---

## 2. Bảng phân loại chi tiết các yêu cầu trong hệ thống Clinic-Cloud

| Tiêu chí | Yêu cầu Chức năng (Functional) | Yêu cầu Phi chức năng (Non-Functional) |
| :--- | :--- | :--- |
| **Định nghĩa** | Mô tả các hành động nghiệp vụ cụ thể | Mô tả thuộc tính chất lượng và ràng buộc kỹ thuật |
| **Ví dụ 1** | Cho phép bác sĩ nhập chẩn đoán, triệu chứng và kê đơn thuốc cho bệnh nhân. | Mã hóa toàn bộ dữ liệu hồ sơ bệnh nhân bằng thuật toán **AES-256** trước khi lưu vào CSDL. |
| **Ví dụ 2** | Cho phép nhân viên lễ tân tìm kiếm bệnh nhân theo Mã bệnh nhân (MBN) hoặc CCCD. | Thời gian phản hồi của hệ thống khi truy vấn hồ sơ không được vượt quá **1.5 giây**. |
| **Ví dụ 3** | Xuất hóa đơn viện phí và đơn thuốc tự động ra file PDF để in. | Tuân thủ tiêu chuẩn bảo mật y tế **HIPAA/GDPR** và ghi nhật ký truy cập (Audit Log) 24/7. |

---

# Phần 2. Triển Khai Bộ Hồ Sơ Hệ Thống Clinic-Cloud

## 1. File Word Đặc Tả SRS (`SRS_Benh_An_Dien_Tu.docx`)
* **Cấu trúc Heading Styles chuẩn:**
  - `Heading 1`: Các chương chính (Giới thiệu, Yêu cầu chức năng, Yêu cầu phi chức năng).
  - `Heading 2`: Các mục chức năng và quy định bảo mật chi tiết.
  - `Normal`: Nội dung mô tả văn bản.
* **Tích hợp Auto TOC:** Trường mục lục tự động `Table of Contents` đặt ở Trang 1, tự động thu thập tiêu đề và cập nhật số trang thời gian thực.

---

## 2. File Excel Quản Lý Bệnh Nhân (`Quan_Ly_Benh_Nhan_Clinic.xlsx`)
* **Tuân thủ 5 Quy tắc Vàng Database Mindset:**
  1. Hàng 1 là **Header Row duy nhất** chứa các tên trường (`Ma_Benh_Nhan`, `Ho_Ten`, `So_Dien_Thoai`, `Email`, `Ngay_Kham`, `Chuan_Doan`, `Trang_Thai_Thanh_Toan`).
  2. Dữ liệu nguyên tử (Atomic Data): Phân tách thông tin liên hệ và chẩn đoán thành từng cột riêng biệt.
  3. Không chèn dòng tính tổng hay tiêu đề trang trí làm ngắt quãng bảng dữ liệu.
* **Xử lý bẫy hàng nghìn ô trống (Missing Data):**
  - *Hiện tượng:* Bệnh nhân không điền Số điện thoại hoặc Email khiến cột dữ liệu có nhiều ô trống (Blank Cells). Nếu bấm AutoFilter thông thường, Excel có thể bị dừng quét vùng dữ liệu tại dòng trống đầu tiên.
  - *Giải pháp:* **Bôi đen chủ động toàn bộ vùng dữ liệu** (từ `A1` đến cột cuối cùng `G1001`) trước khi bấm `Ctrl + Shift + L` để kích hoạt AutoFilter. Việc chọn vùng quét cố định giúp bộ lọc AutoFilter bao phủ 100% dữ liệu mà không bị ngắt quãng giữa chừng bởi ô trống.

---

## 3. Slide Báo Cáo Dịch Bệnh (`Gioi_Thieu_Dich_Benh_Clinic.pptx`)
* **Trực quan hóa tối giản:** Sử dụng **Biểu đồ bánh khuyết (Donut Chart)** để biểu diễn tỷ lệ cơ cấu các ca bệnh (Sốt siêu vi, Viêm họng, Cúm A) thay cho bảng biểu chữ rườm rà.
* **Áp dụng Assertion Title:**
  > *"Sốt siêu vi và Cúm A chiếm 65% tổng số ca khám trong tuần cao điểm tháng 10"*

---

# Phần 3. Báo Cáo Kiểm Thử Gỡ Lỗi Heading Styles (`Ctrl + Shift + N`)

## 1. Phân tích sự cố gán nhầm nhãn Style
Trong quá trình soạn thảo file Word SRS, người dùng vô tình chọn nguyên một đoạn văn mô tả dài 5 trang và click chọn nhãn **Heading 1** (hoặc Heading 2).

**Hậu quả:** 
- Toàn bộ 5 trang văn bản bị coi là "Tiêu đề lớn".
- Khi bấm cập nhật mục lục tự động (`Update Field` Auto TOC), toàn bộ 5 trang văn bản thô bị tràn vào bảng mục lục, làm hỏng hoàn toàn giao diện tài liệu SRS.

---

## 2. Quy trình gỡ lỗi an toàn bằng phím tắt `Ctrl + Shift + N`

1. **Bước 1 (Khoanh vùng lỗi):** Mở thẻ `View` ➔ Tích chọn `Navigation Pane` để xem cây tiêu đề. Phát hiện đoạn văn bản 5 trang đang bị nằm trong danh sách tiêu đề.
2. **Bước 2 (Bôi đen đoạn văn bị lỗi):** Bôi đen toàn bộ khối văn bản 5 trang đang bị gán nhầm nhãn Heading.
3. **Bước 3 (Thực hiện trả nhãn chuẩn):** Nhấn tổ hợp phím tắt **`Ctrl + Shift + N`** trên bàn phím.
   - *Cơ chế kỹ thuật:* Phím tắt `Ctrl + Shift + N` lập tức hủy bỏ nhãn Heading và gán lại nhãn văn bản thường **Normal Style** cho đoạn văn được chọn.
4. **Bước 4 (Cập nhật lại Auto TOC):** Nhấp chuột phải vào bảng Mục lục tự động ➔ Chọn `Update Field` ➔ Chọn `Update entire table`. Đoạn văn 5 trang sẽ ngay lập tức biến mất khỏi bảng mục lục, trả lại cấu trúc tài liệu chuẩn xác 100%.

---

# Phần 4. Danh Sách Kiểm Tra Hoàn Thiện (Task List)

- [x] Phân tích Yêu cầu Chức năng vs Phi chức năng và bảng so sánh trong `README.md`
- [x] Lập quy trình gỡ lỗi gán nhầm Heading bằng `Ctrl + Shift + N` trong `README.md`
- [x] Tạo file Word `SRS_Benh_An_Dien_Tu.docx` chuẩn Heading Styles & Auto TOC
- [x] Tạo file Excel `Quan_Ly_Benh_Nhan_Clinic.xlsx` chuẩn Database và xử lý ô trống AutoFilter[cite: 1]
- [x] Thiết kế Slide `Gioi_Thieu_Dich_Benh_Clinic.pptx` biểu diễn Donut Chart tối giản[cite: 1]
- [ ] Push toàn bộ sản phẩm lên GitHub Repository `HCMK26CNTT2_IT108_Session10_Ex04`
- [ ] Copy đường dẫn link nộp bài dán lên LMS Portal[cite: 1]
