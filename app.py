import os
import io
import tempfile
import threading
import shutil
from flask import Flask, render_template, request, send_file, jsonify, make_response
import webview
from docx_processor import format_document
import logging

# Cấu hình tracking người dùng sử dụng tool
logger = logging.getLogger('autoword_tracker')
logger.setLevel(logging.INFO)
log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'autoword_access.log')
file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
formatter = logging.Formatter('%(asctime)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def log_format_event(req, files, success, error_msg=None):
    try:
        ip = req.headers.get('X-Forwarded-For', req.remote_addr)
        if ip and ',' in ip:
            ip = ip.split(',')[0].strip()
        file_names = [f.filename for f in files if f.filename != '']
        num_files = len(file_names)
        files_str = ", ".join(file_names)
        status_str = "SUCCESS" if success else f"FAILED ({error_msg})"
        logger.info(f"IP: {ip} | Status: {status_str} | Count: {num_files} | Files: [{files_str}]")
    except Exception:
        pass

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Giới hạn upload tối đa 100MB


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/sitemap.xml')
def sitemap():
    return send_file(os.path.join(app.static_folder, 'sitemap.xml'), mimetype='application/xml')


@app.route('/favicon.ico')
def favicon():
    return send_file(os.path.join(app.static_folder, 'logo.png'), mimetype='image/png')


@app.route('/api/feedback', methods=['POST'])
def api_feedback():
    try:
        feedback_type = request.form.get('type', 'Khác')
        message = request.form.get('message', '').strip()
        if not message:
            return jsonify({'error': 'Vui lòng nhập nội dung góp ý.'}), 400
            
        ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if ip and ',' in ip:
            ip = ip.split(',')[0].strip()
            
        import csv
        from datetime import datetime
        feedback_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'feedbacks.csv')
        
        file_exists = os.path.isfile(feedback_file)
        with open(feedback_file, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(['Thời gian', 'IP', 'Loại góp ý', 'Nội dung'])
            writer.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ip, feedback_type, message])
            
        return jsonify({'success': True, 'message': 'Cảm ơn bạn đã gửi góp ý!'})
    except Exception as e:
        return jsonify({'error': f'Lỗi khi gửi góp ý: {str(e)}'}), 500


@app.route('/api/format', methods=['POST'])
def api_format():
    # Kiểm tra file
    if 'file' not in request.files:
        return jsonify({'error': 'Không tìm thấy tệp tin được tải lên.'}), 400

    files = request.files.getlist('file')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'Tên tệp tin không hợp lệ.'}), 400

    for file in files:
        if file.filename != '' and not file.filename.lower().endswith('.docx'):
            return jsonify({'error': 'Vui lòng chỉ tải lên tệp tin Word (.docx).'}), 400

    # Đọc cấu hình định dạng
    try:
        opts = {
            'font_name':            request.form.get('font_name', 'Times New Roman'),
            'body_size':            float(request.form.get('body_size', 14)),
            'line_spacing':         float(request.form.get('line_spacing', 1.5)),
            'space_before':         float(request.form.get('space_before', 0)),
            'space_after':          float(request.form.get('space_after', 0)),
            'first_line_indent':    float(request.form.get('first_line_indent', 10)),
            'margin_top':           float(request.form.get('margin_top', 20)),
            'margin_bottom':        float(request.form.get('margin_bottom', 20)),
            'margin_left':          float(request.form.get('margin_left', 30)),
            'margin_right':         float(request.form.get('margin_right', 15)),
            'heading1_size':        float(request.form.get('heading1_size', 14)),
            'heading1_bold':        request.form.get('heading1_bold') == 'true',
            'heading1_uppercase':   request.form.get('heading1_uppercase') == 'true',
            'heading2_size':        float(request.form.get('heading2_size', 14)),
            'heading2_bold':        request.form.get('heading2_bold') == 'true',
            'heading3_size':        float(request.form.get('heading3_size', 13)),
            'heading3_italic':      request.form.get('heading3_italic') == 'true',
            'auto_number_headings': request.form.get('auto_number_headings') == 'true',
            'format_admin_parts':   request.form.get('format_admin_parts') == 'true',
            'add_page_numbers':     request.form.get('add_page_numbers') == 'true',
            'format_cover':         request.form.get('format_cover') == 'true',
            'alignment':            request.form.get('alignment', 'justify'),
            'contextual_spacing':   request.form.get('contextual_spacing') == 'true',
        }
    except Exception as e:
        return jsonify({'error': f'Dữ liệu cấu hình không đúng: {str(e)}'}), 400

    save_path = request.form.get('save_path', '').strip()
    valid_files = [f for f in files if f.filename != '']
    if not valid_files:
        return jsonify({'error': 'Không có tệp tin hợp lệ nào được tải lên.'}), 400

    temp_dir = tempfile.gettempdir()

    try:
        # TH1: Chạy offline (có save_path)
        if save_path:
            if len(valid_files) == 1:
                # 1 file duy nhất: save_path là đường dẫn tệp tin cụ thể
                file = valid_files[0]
                input_path = os.path.join(temp_dir, 'autoword_input_' + os.path.basename(file.filename))
                file.save(input_path)
                try:
                    stats = format_document(input_path, save_path, opts)
                finally:
                    if os.path.exists(input_path):
                        os.remove(input_path)
                log_format_event(request, valid_files, True)
                return jsonify({'success': True, 'saved_to': save_path, 'stats': stats})
            else:
                # Nhiều file: save_path là một thư mục lưu trữ
                if not os.path.isdir(save_path):
                    # Thử tạo thư mục nếu chưa tồn tại
                    os.makedirs(save_path, exist_ok=True)
                
                saved_files = []
                for file in valid_files:
                    # Tạo file tạm
                    input_path = os.path.join(temp_dir, 'autoword_input_' + os.path.basename(file.filename))
                    file.save(input_path)
                    
                    # Xác định tên file đầu ra trong thư mục đích
                    base, ext = os.path.splitext(os.path.basename(file.filename))
                    output_name = f"{base}_formatted{ext}"
                    dest_path = os.path.join(save_path, output_name)
                    
                    try:
                        format_document(input_path, dest_path, opts)
                        saved_files.append(dest_path)
                    finally:
                        if os.path.exists(input_path):
                            os.remove(input_path)
                log_format_event(request, valid_files, True)
                return jsonify({'success': True, 'saved_to': save_path, 'files': saved_files})
                
        # TH2: Chạy online (tải về qua trình duyệt - không có save_path)
        else:
            if len(valid_files) == 1:
                # 1 file duy nhất: trả về file .docx trực tiếp
                file = valid_files[0]
                input_path = os.path.join(temp_dir, 'autoword_input_' + os.path.basename(file.filename))
                file.save(input_path)
                
                base, _ = os.path.splitext(file.filename)
                output_name = f"{base}_formatted.docx"
                output_path = os.path.join(temp_dir, output_name)
                
                import json
                try:
                    stats = format_document(input_path, output_path, opts)
                    # Load vào RAM để xoá file rác ngay lập tức
                    return_data = io.BytesIO()
                    with open(output_path, 'rb') as f:
                        return_data.write(f.read())
                    return_data.seek(0)
                finally:
                    if os.path.exists(output_path):
                        os.remove(output_path)
                    if os.path.exists(input_path):
                        os.remove(input_path)
                        
                log_format_event(request, valid_files, True)
                response = make_response(send_file(
                    return_data,
                    as_attachment=True,
                    download_name=output_name,
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                ))
                response.headers['X-Format-Stats'] = json.dumps(stats)
                return response
            else:
                # Nhiều file: Đóng gói vào file ZIP
                import zipfile
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for file in valid_files:
                        input_path = os.path.join(temp_dir, 'autoword_input_' + os.path.basename(file.filename))
                        file.save(input_path)
                        
                        base, ext = os.path.splitext(os.path.basename(file.filename))
                        output_name = f"{base}_formatted{ext}"
                        output_path = os.path.join(temp_dir, output_name)
                        
                        try:
                            format_document(input_path, output_path, opts)
                            zip_file.write(output_path, output_name)
                        finally:
                            if os.path.exists(output_path):
                                os.remove(output_path)
                            if os.path.exists(input_path):
                                os.remove(input_path)
                                
                zip_buffer.seek(0)
                log_format_event(request, valid_files, True)
                return send_file(
                    zip_buffer,
                    as_attachment=True,
                    download_name="AutoWord_Formatted_Files.zip",
                    mimetype='application/zip'
                )
    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = str(e)
        if "Permission denied" in err_msg or "PermissionError" in err_msg:
            friendly_msg = "Không thể lưu tệp tin. Có thể tệp tin đích đang được mở trong Word hoặc ứng dụng khác. Vui lòng đóng tệp tin đó lại hoặc chọn nơi lưu khác và thử lại."
        else:
            friendly_msg = f"Lỗi định dạng tài liệu: {err_msg}"
        log_format_event(request, valid_files, False, err_msg)
        return jsonify({'error': friendly_msg}), 500


# ---------------------------------------------------------------------------
# Pywebview JS API Bridge — cho phép JS gọi hàm Python để mở Save Dialog
# ---------------------------------------------------------------------------
class Api:
    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def select_save_path(self, default_name):
        """Mở hộp thoại Save File của Windows, trả về đường dẫn hoặc None."""
        if not self._window:
            return None
        file_types = ('Word Documents (*.docx)',)
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=default_name,
            file_types=file_types
        )
        if result:
            if isinstance(result, (tuple, list)):
                return result[0]
            return result
        return None

    def select_save_folder(self):
        """Mở hộp thoại chọn thư mục của Windows, trả về đường dẫn hoặc None."""
        if not self._window:
            return None
        result = self._window.create_file_dialog(
            webview.FileDialog.FOLDER
        )
        if result:
            if isinstance(result, (tuple, list)):
                return result[0]
            return result
        return None

    def download_cover(self, filename):
        """Mở hộp thoại để tải xuống mẫu bìa tương ứng trong pywebview."""
        if not self._window:
            return {'success': False, 'error': 'Không tìm thấy cửa sổ.'}
        
        display_names = {
            'mau_bia_1.docx': 'Mau_bia_1.docx',
            'mau_bia_2.docx': 'Mau_bia_2.docx',
            'mau_bia_3.docx': 'Mau_bia_3.docx'
        }
        default_name = display_names.get(filename, filename)
        
        save_path = self.select_save_path(default_name)
        if not save_path:
            return {'success': False, 'info': 'Đã hủy tải xuống.'}
            
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            src_path = os.path.join(base_dir, 'static', 'covers', filename)
            
            if not os.path.exists(src_path):
                return {'success': False, 'error': f'Không tìm thấy tệp mẫu {filename}.'}
                
            shutil.copy2(src_path, save_path)
            return {'success': True, 'saved_to': save_path}
        except Exception as e:
            return {'success': False, 'error': f'Lỗi lưu tệp: {str(e)}'}


# ---------------------------------------------------------------------------
# Khởi chạy ứng dụng Desktop
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 50)
    print("  Auto-Word Formatter - Desktop Offline")
    print("=" * 50)

    api = Api()

    window = webview.create_window(
        title='Tool Word By Van Dam (Ủng hộ tôi 2562207069 BIDV)',
        url=app,
        js_api=api,
        width=1200,
        height=850,
        min_size=(960, 700),
        resizable=True,
    )
    api.set_window(window)
    webview.start()
