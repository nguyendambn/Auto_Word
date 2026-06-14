#!/bin/bash
# Script dọn rác tự động cho Auto-Word
# Xóa các file .docx và .zip trong thư mục tạm cũ hơn 1 ngày

# Nếu triển khai bằng Docker, bạn có thể thiết lập script này chạy trên Host Ubuntu
# Hoặc thiết lập bên trong Container (nếu container không tự bị xóa).
find /tmp -name "*.docx" -type f -mmin +1440 -delete
find /tmp -name "*.zip" -type f -mmin +1440 -delete
find /tmp -name "autoword_*" -type f -mmin +1440 -delete

echo "Đã dọn dẹp các file rác cũ hơn 1 ngày trong /tmp."
