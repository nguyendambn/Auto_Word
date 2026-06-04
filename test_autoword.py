"""
Auto-Word Formatter — Bộ kiểm thử tự động
Chạy: python test_autoword.py
"""
import os
import sys
import io
import json
import time
import threading
import tempfile
import shutil
import requests

# ─────────────────────────────────────────
# Cấu hình
# ─────────────────────────────────────────
BASE_URL = "http://127.0.0.1:5050"
SAMPLE_DOCX_SMALL  = r"C:\Users\The Linh\Desktop\testing4.docx"
SAMPLE_DOCX_MEDIUM = r"C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx"
SAMPLE_DOCX_LARGE  = r"C:\Users\The Linh\Desktop\testing1.docx"
OUTPUT_DIR = tempfile.mkdtemp(prefix="autoword_test_")

# ─────────────────────────────────────────
# Helper
# ─────────────────────────────────────────
RESULTS = []
PASS = "✅ PASS"
FAIL = "❌ FAIL"
SKIP = "⏭️  SKIP"

def record(tc_id, name, status, note=""):
    icon = "✅" if status == "PASS" else ("❌" if status == "FAIL" else "⏭️ ")
    RESULTS.append({"id": tc_id, "name": name, "status": status, "note": note})
    print(f"  {icon} [{tc_id}] {name}" + (f" — {note}" if note else ""))

def default_opts():
    return {
        "font_name": "Times New Roman",
        "body_size": "14",
        "line_spacing": "1.5",
        "space_before": "0",
        "space_after": "0",
        "first_line_indent": "10",
        "margin_top": "25",
        "margin_bottom": "20",
        "margin_left": "35",
        "margin_right": "20",
        "heading1_size": "14",
        "heading1_bold": "true",
        "heading1_uppercase": "true",
        "heading2_size": "14",
        "heading2_bold": "true",
        "heading3_size": "13",
        "heading3_italic": "false",
        "auto_number_headings": "true",
        "format_admin_parts": "false",
        "add_page_numbers": "true",
        "format_cover": "false",
        "alignment": "justify",
        "contextual_spacing": "true",
    }

def post_format(docx_path, opts=None, save_path=None):
    """Gửi POST /api/format, trả về (response, elapsed_sec)."""
    if opts is None:
        opts = default_opts()
    data = dict(opts)
    if save_path:
        data["save_path"] = save_path
    with open(docx_path, "rb") as f:
        files = {"file": (os.path.basename(docx_path), f,
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        t0 = time.time()
        r = requests.post(f"{BASE_URL}/api/format", data=data, files=files, timeout=120)
    return r, round(time.time() - t0, 2)

def save_output(resp, name):
    """Lưu blob response ra file tạm, trả về path."""
    out = os.path.join(OUTPUT_DIR, name)
    with open(out, "wb") as f:
        f.write(resp.content)
    return out

# ─────────────────────────────────────────
# Khởi động Flask server nền
# ─────────────────────────────────────────
def start_server():
    """Chạy Flask app ở port 5050 trong background thread."""
    # Patch env để app.py biết chạy server mode (không mở pywebview)
    os.environ["SERVER_MODE"] = "1"

    # Thêm thư mục dự án vào sys.path
    proj = r"C:\Users\The Linh\Desktop\Auto_Word"
    if proj not in sys.path:
        sys.path.insert(0, proj)

    from app import app as flask_app
    # Tắt reloader / debug để tránh fork
    flask_app.run(host="127.0.0.1", port=5050, debug=False, use_reloader=False)

print("🚀 Khởi động Flask server test trên port 5050...")
t = threading.Thread(target=start_server, daemon=True)
t.start()
time.sleep(3)   # Chờ server warm-up

# Kiểm tra server đã lên chưa
try:
    requests.get(f"{BASE_URL}/", timeout=5)
    print("✅ Server sẵn sàng!\n")
except Exception:
    print("❌ Không kết nối được server — dừng kiểm thử")
    sys.exit(1)

# ═══════════════════════════════════════════
# TC-01: Upload & Validate File
# ═══════════════════════════════════════════
print("=" * 60)
print("  TC-01: Upload & Validate File")
print("=" * 60)

# TC-01-01: Upload file .docx hợp lệ (nhỏ)
try:
    r, elapsed = post_format(SAMPLE_DOCX_SMALL)
    ct = r.headers.get("Content-Type","")
    if r.status_code == 200 and ("wordprocessingml" in ct or len(r.content) > 1000):
        record("TC-01-01", "Upload file .docx hợp lệ (nhỏ)", "PASS", f"{elapsed}s — {len(r.content)//1024}KB output")
    else:
        record("TC-01-01", "Upload file .docx hợp lệ (nhỏ)", "FAIL",
               f"status={r.status_code}, ct={ct}")
except Exception as e:
    record("TC-01-01", "Upload file .docx hợp lệ (nhỏ)", "FAIL", str(e))

# TC-01-02: Không gửi file
try:
    r = requests.post(f"{BASE_URL}/api/format", data=default_opts(), timeout=10)
    body = r.json()
    if r.status_code == 400 and "error" in body:
        record("TC-01-02", "Không gửi file (thiếu field 'file')", "PASS", body["error"])
    else:
        record("TC-01-02", "Không gửi file (thiếu field 'file')", "FAIL",
               f"status={r.status_code}")
except Exception as e:
    record("TC-01-02", "Không gửi file (thiếu field 'file')", "FAIL", str(e))

# TC-01-03: Upload file không phải .docx (.txt)
try:
    fake = io.BytesIO(b"Hello world")
    r = requests.post(f"{BASE_URL}/api/format",
                      data=default_opts(),
                      files={"file": ("test.txt", fake, "text/plain")},
                      timeout=10)
    body = r.json()
    if r.status_code == 400 and "docx" in body.get("error","").lower():
        record("TC-01-03", "Upload file không phải .docx", "PASS", body["error"])
    else:
        record("TC-01-03", "Upload file không phải .docx", "FAIL",
               f"status={r.status_code}, body={body}")
except Exception as e:
    record("TC-01-03", "Upload file không phải .docx", "FAIL", str(e))

# TC-01-04: Upload file rỗng (0 byte)
try:
    empty = io.BytesIO(b"")
    r = requests.post(f"{BASE_URL}/api/format",
                      data=default_opts(),
                      files={"file": ("empty.docx", empty,
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                      timeout=10)
    # Kỳ vọng: lỗi (400 hoặc 500), không crash server
    if r.status_code in (400, 500):
        record("TC-01-04", "Upload file rỗng 0 byte", "PASS",
               f"status={r.status_code} — server không crash")
    else:
        record("TC-01-04", "Upload file rỗng 0 byte", "FAIL",
               f"status={r.status_code} (kỳ vọng 400/500)")
except Exception as e:
    record("TC-01-04", "Upload file rỗng 0 byte", "FAIL", str(e))

# TC-01-06: Upload file lớn (> 5MB)
try:
    r, elapsed = post_format(SAMPLE_DOCX_MEDIUM)
    ct = r.headers.get("Content-Type","")
    if r.status_code == 200 and ("wordprocessingml" in ct or len(r.content) > 10000):
        size_mb = os.path.getsize(SAMPLE_DOCX_MEDIUM)//1024//1024
        record("TC-01-06", f"Upload file lớn ({size_mb}MB)",
               "PASS", f"{elapsed}s — output {len(r.content)//1024}KB")
    else:
        ct = r.headers.get("Content-Type","?")
        snippet = r.text[:200] if "json" in ct else f"Binary ({len(r.content)} bytes)"
        record("TC-01-06", "Upload file lớn (>5MB)", "FAIL",
               f"status={r.status_code}, body={snippet}")
except Exception as e:
    record("TC-01-06", "Upload file lớn (>5MB)", "FAIL", str(e))

# TC-01-07: Upload file .docx giả (nội dung không phải OOXML)
try:
    fake_docx = io.BytesIO(b"PK\x03\x04" + b"\x00" * 100)  # ZIP header nhưng rỗng
    r = requests.post(f"{BASE_URL}/api/format",
                      data=default_opts(),
                      files={"file": ("fake.docx", fake_docx,
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                      timeout=15)
    if r.status_code in (400, 500):
        record("TC-01-07", "Upload file .docx giả (corrupt)", "PASS",
               f"status={r.status_code} — server trả lỗi đúng")
    else:
        record("TC-01-07", "Upload file .docx giả (corrupt)", "FAIL",
               f"status={r.status_code} (kỳ vọng lỗi)")
except Exception as e:
    record("TC-01-07", "Upload file .docx giả (corrupt)", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-02: Cấu hình Định dạng Cơ bản
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-02: Cấu hình Định dạng Cơ bản")
print("=" * 60)

# TC-02-01: Preset mặc định
try:
    r, elapsed = post_format(SAMPLE_DOCX_SMALL)
    if r.status_code == 200:
        record("TC-02-01", "Preset mặc định HaUI", "PASS", f"{elapsed}s")
        save_output(r, "tc0201_default.docx")
    else:
        record("TC-02-01", "Preset mặc định HaUI", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-01", "Preset mặc định HaUI", "FAIL", str(e))

# TC-02-02: Thay đổi font chữ → Arial
try:
    opts = default_opts()
    opts["font_name"] = "Arial"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-02", "Font chữ Arial", "PASS", f"{elapsed}s")
    else:
        record("TC-02-02", "Font chữ Arial", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-02", "Font chữ Arial", "FAIL", str(e))

# TC-02-03: Cỡ chữ 12pt
try:
    opts = default_opts()
    opts["body_size"] = "12"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-03", "Cỡ chữ 12pt", "PASS", f"{elapsed}s")
    else:
        record("TC-02-03", "Cỡ chữ 12pt", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-03", "Cỡ chữ 12pt", "FAIL", str(e))

# TC-02-04: Dãn dòng 2.0
try:
    opts = default_opts()
    opts["line_spacing"] = "2.0"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-04", "Dãn dòng 2.0", "PASS", f"{elapsed}s")
    else:
        record("TC-02-04", "Dãn dòng 2.0", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-04", "Dãn dòng 2.0", "FAIL", str(e))

# TC-02-05: Lề trái 40mm, phải 10mm
try:
    opts = default_opts()
    opts["margin_left"] = "40"
    opts["margin_right"] = "10"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-05", "Lề trái 40mm, phải 10mm", "PASS", f"{elapsed}s")
    else:
        record("TC-02-05", "Lề trái 40mm, phải 10mm", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-05", "Lề trái 40mm, phải 10mm", "FAIL", str(e))

# TC-02-06: Thụt đầu dòng 0mm
try:
    opts = default_opts()
    opts["first_line_indent"] = "0"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-06", "Thụt đầu dòng 0mm (tắt)", "PASS", f"{elapsed}s")
    else:
        record("TC-02-06", "Thụt đầu dòng 0mm (tắt)", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-06", "Thụt đầu dòng 0mm (tắt)", "FAIL", str(e))

# TC-02-07: Căn lề Left
try:
    opts = default_opts()
    opts["alignment"] = "left"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-02-07", "Căn lề Left", "PASS", f"{elapsed}s")
    else:
        record("TC-02-07", "Căn lề Left", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-02-07", "Căn lề Left", "FAIL", str(e))

# TC-02-08: Input giá trị chữ (không hợp lệ)
try:
    opts = default_opts()
    opts["body_size"] = "abc"
    r = requests.post(f"{BASE_URL}/api/format",
                      data=opts,
                      files={"file": (os.path.basename(SAMPLE_DOCX_SMALL),
                             open(SAMPLE_DOCX_SMALL, "rb"),
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                      timeout=15)
    if r.status_code == 400:
        record("TC-02-08", "body_size='abc' → lỗi 400", "PASS", r.json().get("error",""))
    else:
        record("TC-02-08", "body_size='abc' → lỗi 400", "FAIL",
               f"status={r.status_code} (kỳ vọng 400)")
except Exception as e:
    record("TC-02-08", "body_size='abc' → lỗi 400", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-03: Heading & Auto-numbering
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-03: Heading & Đánh số mục tự động")
print("=" * 60)

# TC-03-01: Auto-number heading bật
try:
    opts = default_opts()
    opts["auto_number_headings"] = "true"
    r, elapsed = post_format(SAMPLE_DOCX_MEDIUM, opts)
    if r.status_code == 200:
        record("TC-03-01", "Auto-number headings = ON", "PASS", f"{elapsed}s")
        save_output(r, "tc0301_heading_auto.docx")
    else:
        record("TC-03-01", "Auto-number headings = ON", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-03-01", "Auto-number headings = ON", "FAIL", str(e))

# TC-03-02: Auto-number heading tắt
try:
    opts = default_opts()
    opts["auto_number_headings"] = "false"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-03-02", "Auto-number headings = OFF", "PASS", f"{elapsed}s")
    else:
        record("TC-03-02", "Auto-number headings = OFF", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-03-02", "Auto-number headings = OFF", "FAIL", str(e))

# TC-03-03: Heading 1 uppercase + bold
try:
    opts = default_opts()
    opts["heading1_bold"] = "true"
    opts["heading1_uppercase"] = "true"
    opts["heading1_size"] = "14"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-03-03", "Heading 1 bold + uppercase", "PASS", f"{elapsed}s")
    else:
        record("TC-03-03", "Heading 1 bold + uppercase", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-03-03", "Heading 1 bold + uppercase", "FAIL", str(e))

# TC-03-04: Heading 3 italic
try:
    opts = default_opts()
    opts["heading3_italic"] = "true"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-03-04", "Heading 3 italic = ON", "PASS", f"{elapsed}s")
    else:
        record("TC-03-04", "Heading 3 italic = ON", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-03-04", "Heading 3 italic = ON", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-04: Đánh số trang
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-04: Đánh số trang tự động")
print("=" * 60)

# TC-04-01: Bật đánh số trang
try:
    opts = default_opts()
    opts["add_page_numbers"] = "true"
    r, elapsed = post_format(SAMPLE_DOCX_MEDIUM, opts)
    if r.status_code == 200:
        record("TC-04-01", "Đánh số trang = ON", "PASS", f"{elapsed}s")
        save_output(r, "tc0401_page_numbers.docx")
    else:
        record("TC-04-01", "Đánh số trang = ON", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-04-01", "Đánh số trang = ON", "FAIL", str(e))

# TC-04-02: Tắt đánh số trang
try:
    opts = default_opts()
    opts["add_page_numbers"] = "false"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-04-02", "Đánh số trang = OFF", "PASS", f"{elapsed}s")
    else:
        record("TC-04-02", "Đánh số trang = OFF", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-04-02", "Đánh số trang = OFF", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-05: Định dạng Trang bìa
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-05: Định dạng Trang bìa")
print("=" * 60)

# TC-05-01: Bật format_cover
try:
    opts = default_opts()
    opts["format_cover"] = "true"
    r, elapsed = post_format(SAMPLE_DOCX_MEDIUM, opts)
    if r.status_code == 200:
        record("TC-05-01", "format_cover = ON", "PASS", f"{elapsed}s")
        save_output(r, "tc0501_cover.docx")
    else:
        record("TC-05-01", "format_cover = ON", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-05-01", "format_cover = ON", "FAIL", str(e))

# TC-05-02: Tắt format_cover
try:
    opts = default_opts()
    opts["format_cover"] = "false"
    r, elapsed = post_format(SAMPLE_DOCX_SMALL, opts)
    if r.status_code == 200:
        record("TC-05-02", "format_cover = OFF", "PASS", f"{elapsed}s")
    else:
        record("TC-05-02", "format_cover = OFF", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-05-02", "format_cover = OFF", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-06: Tất cả tùy chọn bật cùng lúc
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-06: Kiểm tra toàn diện — Bật tất cả tùy chọn")
print("=" * 60)

try:
    opts = {
        "font_name": "Times New Roman",
        "body_size": "14",
        "line_spacing": "1.5",
        "space_before": "0",
        "space_after": "0",
        "first_line_indent": "10",
        "margin_top": "25",
        "margin_bottom": "20",
        "margin_left": "35",
        "margin_right": "20",
        "heading1_size": "14",
        "heading1_bold": "true",
        "heading1_uppercase": "true",
        "heading2_size": "14",
        "heading2_bold": "true",
        "heading3_size": "13",
        "heading3_italic": "false",
        "auto_number_headings": "true",
        "format_admin_parts": "true",
        "add_page_numbers": "true",
        "format_cover": "true",
        "alignment": "justify",
        "contextual_spacing": "true",
    }
    r, elapsed = post_format(SAMPLE_DOCX_MEDIUM, opts)
    if r.status_code == 200:
        out = save_output(r, "tc06_full_options.docx")
        size_kb = os.path.getsize(out) // 1024
        record("TC-06-01", "Tất cả tùy chọn bật đồng thời", "PASS",
               f"{elapsed}s — output {size_kb}KB")
    else:
        ct = r.headers.get("Content-Type","?")
        record("TC-06-01", "Tất cả tùy chọn bật đồng thời", "FAIL",
               f"status={r.status_code}, ct={ct}")
except Exception as e:
    record("TC-06-01", "Tất cả tùy chọn bật đồng thời", "FAIL", str(e))


# ═══════════════════════════════════════════
# TC-08: Xử lý lỗi
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  TC-08: Xử lý lỗi & Giới hạn")
print("=" * 60)

# TC-08-01: GET /api/format (không hỗ trợ)
try:
    r = requests.get(f"{BASE_URL}/api/format", timeout=5)
    if r.status_code == 405:
        record("TC-08-01", "GET /api/format → 405 Method Not Allowed", "PASS")
    else:
        record("TC-08-01", "GET /api/format → 405 Method Not Allowed", "FAIL",
               f"status={r.status_code}")
except Exception as e:
    record("TC-08-01", "GET /api/format", "FAIL", str(e))

# TC-08-02: Kiểm tra trang chủ hoạt động
try:
    r = requests.get(f"{BASE_URL}/", timeout=5)
    if r.status_code == 200 and "html" in r.headers.get("Content-Type",""):
        record("TC-08-02", "GET / → trả về HTML đúng", "PASS")
    else:
        record("TC-08-02", "GET / → trả về HTML đúng", "FAIL",
               f"status={r.status_code}")
except Exception as e:
    record("TC-08-02", "GET /", "FAIL", str(e))

# TC-08-03: File DOCX lớn nhất có thể (~10MB)
try:
    r, elapsed = post_format(SAMPLE_DOCX_LARGE)
    if r.status_code == 200:
        out = save_output(r, "tc0803_very_large.docx")
        size_kb = os.path.getsize(out) // 1024
        record("TC-08-03", f"File rất lớn ({os.path.getsize(SAMPLE_DOCX_LARGE)//1024//1024}MB)", 
               "PASS", f"{elapsed}s — output {size_kb}KB")
    else:
        record("TC-08-03", "File rất lớn (~10MB)", "FAIL", f"status={r.status_code}")
except Exception as e:
    record("TC-08-03", "File rất lớn (~10MB)", "FAIL", str(e))


# ═══════════════════════════════════════════
# Tổng kết
# ═══════════════════════════════════════════
print("\n" + "=" * 60)
print("  📊 TỔNG KẾT KẾT QUẢ KIỂM THỬ")
print("=" * 60)

passed = [r for r in RESULTS if r["status"] == "PASS"]
failed = [r for r in RESULTS if r["status"] == "FAIL"]
skipped = [r for r in RESULTS if r["status"] == "SKIP"]

print(f"\n  Tổng số test:  {len(RESULTS)}")
print(f"  ✅ PASS:       {len(passed)}")
print(f"  ❌ FAIL:       {len(failed)}")
print(f"  ⏭️  SKIP:       {len(skipped)}")

if failed:
    print("\n  🔴 Danh sách FAIL:")
    for r in failed:
        print(f"     [{r['id']}] {r['name']}")
        if r["note"]:
            print(f"            → {r['note']}")

print(f"\n  📁 File output lưu tại: {OUTPUT_DIR}")
print("\n✅ Kiểm thử hoàn tất!")
