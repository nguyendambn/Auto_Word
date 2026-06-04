FROM python:3.11-slim

# Thiết lập thư mục làm việc trong container
WORKDIR /app

# Cài đặt các thư viện hệ thống cần thiết (nếu có)
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy và cài đặt các thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn vào container
COPY . .

# Phân quyền cho thư mục /app (Hugging Face chạy với user không có quyền root)
RUN chmod -R 777 /app

# Khai báo biến môi trường chạy chế độ Server
ENV SERVER_MODE=true

# Cổng mặc định của Hugging Face Spaces là 7860
EXPOSE 7860

# Khởi chạy ứng dụng bằng Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "2", "--timeout", "120", "app:app"]
