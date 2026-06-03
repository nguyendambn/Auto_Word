### Các bước cài đặt

1. **Tải mã nguồn về máy hoặc Clone repository**:
   ```bash
   git clone https://github.com/nguyendambn/Auto_Word.git
   cd Auto_Word
   ```

2. **Khởi tạo môi trường ảo (Khuyên dùng)**:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate   # Trên Windows
   source .venv/bin/activate  # Trên macOS/Linux
   ```

3. **Cài đặt các thư viện phụ thuộc**:
   ```bash
   pip install -r requirements.txt
   ```

### Khởi chạy ứng dụng

* **Cách 1: Sử dụng file chạy nhanh (Windows)**:
  Kích đúp chuột vào file `run.bat` ở thư mục gốc của dự án. File này sẽ tự động kích hoạt môi trường ảo và chạy ứng dụng.

* **Cách 2: Chạy trực tiếp qua Terminal**:
  ```bash
  python app.py
  ```
# 📝 Auto-Word Formatter by Van Dam - Công Cụ Định Dạng Báo Cáo & Đồ Án Tốt Nghiệp Tự Động

**Auto-Word Formatter** là một ứng dụng desktop ngoại tuyến (offline) mạnh mẽ được viết bằng **Python** và **JavaScript**, giúp tự động hóa toàn bộ quy trình căn lề, chỉnh sửa font chữ, dãn dòng, tạo đề mục, đánh số trang và định dạng bảng biểu cho các tài liệu Word (`.docx`). 

Công cụ được tối ưu hóa đặc biệt cho đối tượng là sinh viên thực hiện đồ án tốt nghiệp, báo cáo kỹ thuật, luận văn thạc sĩ và nhân viên văn phòng cần định dạng văn bản hành chính chuẩn theo quy định của Đại học công nghiệp Hà Nội

---

## 🌟 Các Chức Năng Chính

### 1. Chuẩn Hóa Lề Trang & Khổ Giấy
* Tự động chuyển đổi khổ giấy sang chuẩn **A4** (210mm x 297mm).
* Thiết lập lề trang linh hoạt theo mm: **Lề trên/dưới** (mặc định 20mm), **Lề trái** (mặc định 30mm) phục vụ đóng gáy sách, **Lề phải** (mặc định 15mm).

### 2. Định Dạng Văn Bản & Dãn Dòng Chuyên Nghiệp
* Đồng nhất phông chữ toàn bộ tài liệu (mặc định **Times New Roman** hoặc phông tùy chọn).
* Căn chỉnh cỡ chữ nội dung (mặc định 14pt), khoảng cách dãn dòng (mặc định 1.5 lines) và khoảng cách giữa các đoạn văn (Space Before/After là 0pt 0pt).
* Tự động thụt lề dòng đầu tiên của các đoạn văn (First Line Indent) và hỗ trợ căn lề đều hai bên (**Justify**) một cách tự nhiên.
* Tự động căn giữa hình ảnh và chú thích ảnh.

### 3. Nhận Diện & Đánh Số Đề Mục Tự Động (Heading Auto-numbering)
* Tự động phân cấp các tiêu đề từ **Heading 1** đến **Heading 9**.
* Thuật toán thông minh tự động nhận diện và gộp các tiêu đề chương bị tách làm 2 dòng (Dòng 1: `CHƯƠNG 1`, Dòng 2: `CƠ SỞ LÝ THUYẾT`) thành một tiêu đề duy nhất chứa dấu ngắt dòng mềm để hiển thị đẹp mắt trong Mục lục.
* Xử lý thông minh các dòng xuống dòng mềm (`Shift+Enter`) chứa tiêu đề và nội dung bị gộp chung đoạn trong tài liệu gốc.
* **Auto-numbering**: Đánh số thứ tự các mục tự động dạng phân cấp thụt lề (Ví dụ: `1.`, `1.1.`, `1.1.1.`, `1.1.1.1.`) và tự động reset bộ đếm khi sang chương mới.
* **Bộ lọc khử nhận diện nhầm**: Loại bỏ các đoạn văn bản thường nhưng định dạng chữ in đậm hoặc bắt đầu bằng số thứ tự giả lập khỏi danh sách tiêu đề.

### 4. Đánh Số Trang Thông Minh (Dual Page Numbering)
* Tự động phân tách tài liệu thành các Section riêng biệt: Trang bìa, Lời cảm ơn, Danh mục viết tắt/hình ảnh, Thân bài và Tài liệu tham khảo.
* **Trang bìa**: Tự động nhận diện trang bìa dựa trên bảng khung viền (Cover Table) và ẩn hoàn toàn số trang.
* **Phần danh mục đầu trang (Front Matter)**: Đánh số trang tự động dạng chữ số La Mã viết thường (`i`, `ii`, `iii`,...) theo đúng chuẩn học thuật.
* **Phần thân bài (Body Text)**: Tự động đánh số trang bắt đầu từ **1** (chữ số Ả Rập) từ phần mở đầu or chương 1  đến hết tài liệu( trừ tài liệu tham khảo).
* Số trang được căn giữa hoàn hảo tại vị trí header hoặc footer theo chiều dọc.


### 6. Căn Chỉnh Bảng Biểu & Chú Thích Tự Động
* Tự động căn giữa tất cả các bảng biểu trong tài liệu.
* Điều chỉnh font chữ, cỡ chữ, căn lề và độ dãn dòng trong các ô của bảng biểu (Table Cells) cho dễ nhìn hơn.
* Nhận diện và tự động đánh số thứ tự chú thích hình ảnh, bảng biểu theo số chương hiện tại (Ví dụ: `Hình 1.1`, `Hình 1.2`, `Bảng 3.1`).

---

## 🛠️ Công Nghệ Sử Dụng

* **Backend**: Python 3.x
  * `python-docx`: Thư viện đọc, chỉnh sửa và ghi cấu trúc XML của tệp tin Word.
  * `lxml`: Xử lý trực tiếp các thẻ định dạng OpenXML của Microsoft Word.
  * `Flask`: Cung cấp API xử lý định dạng tệp tin.
  * `pywebview`: Đóng gói ứng dụng Flask thành phần mềm chạy cửa sổ Desktop chuyên nghiệp.
* **Frontend**: HTML5, CSS3 (Vanilla), JavaScript (ES6+)
  * Giao diện thiết kế theo phong cách hiện đại, hỗ trợ kéo thả tệp tin (Drag and Drop File), hiển thị tiến trình xử lý (Progress Bar).

---

## 💻 Hướng Dẫn Cài Đặt & Sử Dụng

### Yêu cầu hệ thống
* Máy tính đã cài đặt **Python 3.10** trở lên.
* Hệ điều hành: Windows, macOS hoặc Linux.



Sau khi khởi chạy, một cửa sổ Desktop trực quan của công cụ sẽ hiện ra. Bạn chỉ cần chọn file Word cần định dạng, tùy chỉnh các thông số dãn dòng, cỡ chữ trên thanh công cụ bên trái và bấm **Bắt đầu định dạng**.

---

## 📂 Cấu Trúc Mã Nguồn

```text
├── .gitignore               # Cấu hình bỏ qua các tệp không cần thiết khi push Git
├── requirements.txt         # Danh sách thư viện Python phụ thuộc
├── run.bat                  # Script khởi động nhanh ứng dụng trên Windows
├── app.py                   # Điểm khởi chạy chính (Flask API & PyWebview wrapper)
├── docx_processor.py        # Logic cốt lõi xử lý, phân tích và định dạng file Word (.docx)
├── templates/
│   └── index.html           # Giao diện chính của ứng dụng
├── static/
│   ├── css/
│   │   └── style.css        # Định dạng giao diện hiện đại (Glassmorphism, Sidebar layout)
│   └── js/
│   │   └── app.js           # Xử lý các sự kiện click, kéo thả file và gọi API định dạng
└── scratch/                 # Thư mục chứa các script nháp phục vụ phát triển
```

---

## 🔒 Cam Kết Bảo Mật

* Ứng dụng chạy hoàn toàn **Offline** trên máy tính cá nhân của bạn.
* Không có bất kỳ dữ liệu văn bản hay thông tin cá nhân nào được gửi lên internet, đảm bảo tuyệt đối tính bảo mật cho báo cáo tốt nghiệp và văn bản mật của bạn.

---

## 🤝 Hỗ Trợ Dự Án

Nếu công cụ giúp ích được cho đồ án tốt nghiệp hoặc công việc của bạn, hãy ủng hộ tác giả bằng cách Star repository này nhé!
* **Thông tin ủng hộ**: 2562207069 - BIDV (Nguyễn Văn Đàm)
