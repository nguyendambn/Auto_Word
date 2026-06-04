## 💻 Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Tải và thiết lập (Dành cho Developer)

**Bước 1**: Tải mã nguồn về máy hoặc Clone repository:
```bash
git clone https://github.com/nguyendambn/Auto_Word.git
cd Auto_Word
```

**Bước 2**: Khởi tạo môi trường ảo (Khuyên dùng):
```bash
python -m venv .venv
.venv\Scripts\activate   # Trên Windows
source .venv/bin/activate  # Trên macOS/Linux
```

**Bước 3**: Cài đặt các thư viện phụ thuộc:
```bash
pip install -r requirements.txt
```

### 2. Khởi chạy ứng dụng

* **Cách 1: Kích đúp file `.bat` (Windows)**:
  Kích đúp chuột vào file `run.bat` ở thư mục gốc của dự án. Hệ thống sẽ tự động kích hoạt môi trường ảo và mở cửa sổ ứng dụng.
* **Cách 2: Chạy trực tiếp qua Terminal**:
  ```bash
  python app.py
  ```

### 3. Sử dụng
1. Sau khi khởi chạy, giao diện sẽ hiện lên. 
2. **Kéo thả** file Word cần định dạng vào vùng Upload hoặc click để chọn file.
3. Tùy chỉnh các thông số (Phông chữ, Cỡ chữ, Lề, Bật/Tắt đánh số trang tự động...).
4. Bấm **Bắt đầu định dạng**. 
5. Công cụ sẽ xuất ra file có hậu tố `_formatted.docx` hoàn chỉnh.

---
# 📝 Auto-Word Formatter by Van Dam - Công Cụ Định Dạng Báo Cáo & Đồ Án Tốt Nghiệp Tự Động

**Auto-Word Formatter** là một ứng dụng desktop ngoại tuyến (offline) mạnh mẽ được viết bằng **Python** và **JavaScript**, giúp tự động hóa toàn bộ quy trình căn lề, chỉnh sửa font chữ, dãn dòng, tạo đề mục, đánh số trang và định dạng bảng biểu cho các tài liệu Word (`.docx`). 

Công cụ được tối ưu hóa đặc biệt cho đối tượng là sinh viên thực hiện đồ án tốt nghiệp, báo cáo kỹ thuật, luận văn thạc sĩ và nhân viên văn phòng cần định dạng văn bản hành chính chuẩn theo quy định của **Đại học Công nghiệp Hà Nội (HaUI)**.

---

## 🌟 Các Chức Năng Chính Vượt Trội

### 1. Chuẩn Hóa Lề Trang & Khổ Giấy
* Tự động chuyển đổi khổ giấy sang chuẩn **A4** (210mm x 297mm).
* Thiết lập lề trang linh hoạt theo mm: **Lề trên/dưới** (mặc định 20mm), **Lề trái** (mặc định 30mm) phục vụ đóng gáy sách, **Lề phải** (mặc định 15mm).

### 2. Định Dạng Văn Bản & Dãn Dòng Chuyên Nghiệp
* Đồng nhất phông chữ toàn bộ tài liệu (mặc định **Times New Roman** hoặc phông tùy chọn).
* Căn chỉnh cỡ chữ nội dung (mặc định 14pt), khoảng cách dãn dòng (mặc định 1.5 lines) và khoảng cách giữa các đoạn văn (Space Before/After là 0pt 0pt).
* Tự động thụt lề dòng đầu tiên của các đoạn văn (First Line Indent) và hỗ trợ căn lề đều hai bên (**Justify**) một cách tự nhiên.
* Tự động căn giữa hình ảnh và chú thích ảnh, giữ nguyên tỷ lệ thẩm mỹ.
* Hỗ trợ chức năng tắt giãn dòng với các đoạn cùng kiểu (Contextual Spacing).

### 3. Nhận Diện & Đánh Số Đề Mục Tự Động (Heading Auto-numbering)
* Tự động phân cấp các tiêu đề từ **Heading 1** đến **Heading 9**.
* Thuật toán thông minh tự động nhận diện và gộp các tiêu đề chương bị tách làm 2 dòng (Dòng 1: `CHƯƠNG 1`, Dòng 2: `CƠ SỞ LÝ THUYẾT`) thành một tiêu đề duy nhất chứa dấu ngắt dòng mềm để hiển thị đẹp mắt trong Mục lục.
* Hỗ trợ tự động nhận diện **"GIỚI THIỆU ĐỀ TÀI"**, **"MỞ ĐẦU"** là Heading 1 nếu nằm riêng biệt ở đầu trang.
* Xử lý thông minh các dòng xuống dòng mềm (`Shift+Enter`) chứa tiêu đề và nội dung bị gộp chung đoạn trong tài liệu gốc.
* **Auto-numbering**: Đánh số thứ tự các mục tự động dạng phân cấp thụt lề (Ví dụ: `1.`, `1.1.`, `1.1.1.`, `1.1.1.1.`) và tự động reset bộ đếm khi sang chương mới.
* Khử nhận diện nhầm: Loại bỏ các đoạn văn bản thường nhưng định dạng chữ in đậm hoặc bắt đầu bằng số thứ tự giả lập khỏi danh sách tiêu đề. Khử lỗi thụt lề với các danh sách bullet hoặc number chuẩn.

### 4. Đánh Số Trang & Chia Section Thông Minh
* Tự động phân tách tài liệu thành các Section riêng biệt: Trang bìa, Lời cảm ơn, Danh mục viết tắt/hình ảnh, Thân bài, Tài liệu tham khảo, và Phụ lục.
* **Trang bìa**: Tự động nhận diện trang bìa dựa trên bảng khung viền (Cover Table) và ẩn hoàn toàn số trang.
* **Phần danh mục đầu trang (Front Matter)**: Đánh số trang tự động dạng chữ số La Mã viết thường (`i`, `ii`, `iii`,...) theo chuẩn học thuật.
* **Phần thân bài (Body Text)**: Tự động đánh số trang bắt đầu từ **1** (chữ số Ả Rập) từ phần mở đầu / chương 1.
* **Phụ lục & Tài liệu tham khảo**: Tự động nhận diện, tách trang riêng (hỗ trợ cả trang nằm ngang cho Phụ lục lớn) và không đánh số trang cho phần này để tránh sai lệch trang luận văn chính.

### 5. Căn Chỉnh Bảng Biểu & Chú Thích Tự Động
* Tự động căn giữa tất cả các bảng biểu trong tài liệu.
* Điều chỉnh font chữ, cỡ chữ, căn lề và độ dãn dòng trong các ô của bảng biểu (Table Cells) gọn gàng, tránh thừa dòng.
* Tự động bỏ các dòng chú thích rác dưới danh mục bảng biểu và hình ảnh.

### 6. Cung Cấp Sẵn Mẫu Bìa Chuẩn
* Tích hợp sẵn chức năng tải về trực tiếp **3 mẫu bìa chuẩn** trên giao diện:
  * **Mẫu 1**: Dành cho bài tập lớn.
  * **Mẫu 2 & 3**: Dành cho đồ án chuyên ngành / đồ án tốt nghiệp.

---

## 🛠️ Công Nghệ Sử Dụng

* **Backend**: Python 3.x
  * `python-docx`: Thư viện đọc, chỉnh sửa và cấu trúc XML của tệp tin Word.
  * `lxml`: Xử lý trực tiếp các thẻ định dạng OpenXML của Microsoft Word nhằm can thiệp sâu vào số trang và ngắt trang.
  * `Flask`: Cung cấp RESTful API xử lý định dạng tệp tin.
  * `pywebview`: Đóng gói ứng dụng Flask thành phần mềm cửa sổ Desktop chuyên nghiệp.
* **Frontend**: HTML5, CSS3 (Vanilla), JavaScript (ES6+)
  * Giao diện UI/UX hiện đại (Glassmorphism), chế độ Sáng/Tối (Dark/Light mode).
  * Hỗ trợ kéo thả tệp tin (Drag and Drop), hiển thị tiến trình xử lý trực quan (Progress Bar).

---



## 📂 Cấu Trúc Mã Nguồn

```text
Auto_Word/
├── .gitignore               # Cấu hình bỏ qua các tệp không cần thiết (Git)
├── requirements.txt         # Danh sách thư viện Python phụ thuộc
├── run.bat                  # Script khởi động nhanh ứng dụng trên Windows
├── app.py                   # Điểm khởi chạy chính (Flask API & PyWebview wrapper)
├── docx_processor.py        # Module cốt lõi: xử lý thuật toán phân tích và định dạng docx
├── templates/
│   └── index.html           # Giao diện chính của ứng dụng
├── static/
│   ├── covers/              # Chứa các mẫu bìa tải về (mau_bia_1, 2, 3)
│   ├── css/
│   │   └── style.css        # Định dạng giao diện hiện đại (Glassmorphism, Sidebar)
│   ├── js/
│   │   └── app.js           # Xử lý các sự kiện Frontend, kéo thả file, tương tác API
│   └── logo.png             # Logo của ứng dụng
└── scratch/                 # Thư mục nháp (có thể bỏ qua)
```

---

## 🔒 Cam Kết Bảo Mật

* Ứng dụng chạy hoàn toàn **Offline** trên máy tính cá nhân của bạn thông qua webview cục bộ.
* Tuyệt đối **không** có bất kỳ dữ liệu văn bản hay thông tin cá nhân nào được gửi lên internet, đảm bảo tính bảo mật 100% cho báo cáo thực tập, đồ án tốt nghiệp và văn bản mật của cơ quan.

---

## 🤝 Hỗ Trợ Dự Án

Nếu công cụ giúp ích được cho quá trình làm đồ án hoặc công việc của bạn, hãy ủng hộ tác giả bằng cách thả **Star** cho repository này nhé!

* **Chủ sở hữu & Phát triển**: Nguyễn Văn Đàm (Đại học Công nghiệp Hà Nội)
* **Thông tin ủng hộ**: `2562207069` - BIDV (Nguyễn Văn Đàm) ❤️

Cảm ơn bạn đã tin tưởng và sử dụng Auto-Word Formatter!
