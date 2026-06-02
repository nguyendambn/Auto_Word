import os
import io
import tempfile
import threading
from flask import Flask, render_template, request, send_file, jsonify
from docx_processor import format_document

IS_VERCEL = os.environ.get('VERCEL') == '1'
webview = None
if not IS_VERCEL:
    try:
        import webview
    except ImportError:
        pass


app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['TEMPLATES_AUTO_RELOAD'] = True


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/format', methods=['POST'])
def api_format():
    # Kiểm tra file
    if 'file' not in request.files:
        return jsonify({'error': 'Không tìm thấy tệp tin được tải lên.'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Tên tệp tin không hợp lệ.'}), 400

    if not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Vui lòng chỉ tải lên tệp tin Word (.docx).'}), 400

    # Đọc cấu hình định dạng
    try:
        opts = {
            'font_name':            request.form.get('font_name', 'Times New Roman'),
            'body_size':            float(request.form.get('body_size', 13)),
            'line_spacing':         float(request.form.get('line_spacing', 1.3)),
            'space_before':         float(request.form.get('space_before', 0)),
            'space_after':          float(request.form.get('space_after', 6)),
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
            'alignment':            request.form.get('alignment', 'justify'),
            'contextual_spacing':   request.form.get('contextual_spacing') == 'true',
        }
    except Exception as e:
        return jsonify({'error': f'Dữ liệu cấu hình không đúng: {str(e)}'}), 400

    save_path = request.form.get('save_path', '').strip()

    try:
        # Lưu file tạm
        temp_dir = tempfile.gettempdir()
        input_path = os.path.join(temp_dir, 'autoword_input_' + os.path.basename(file.filename))
        file.save(input_path)

        try:
            if save_path:
                # Lưu trực tiếp vào đường dẫn người dùng chọn
                format_document(input_path, save_path, opts)
                return jsonify({'success': True, 'saved_to': save_path})
            else:
                # Tạo file tạm và trả về dưới dạng download
                base, _ = os.path.splitext(file.filename)
                output_name = f"{base}_formatted.docx"
                output_path = os.path.join(temp_dir, output_name)
                
                try:
                    format_document(input_path, output_path, opts)
                    # Load vào RAM để xoá file rác ngay lập tức
                    return_data = io.BytesIO()
                    with open(output_path, 'rb') as f:
                        return_data.write(f.read())
                    return_data.seek(0)
                finally:
                    if os.path.exists(output_path):
                        os.remove(output_path)
                        
                return send_file(
                    return_data,
                    as_attachment=True,
                    download_name=output_name,
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
        finally:
            if os.path.exists(input_path):
                os.remove(input_path)
    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = str(e)
        if "Permission denied" in err_msg or "PermissionError" in err_msg:
            friendly_msg = "Không thể lưu tệp tin. Có thể tệp tin đích đang được mở trong Word hoặc ứng dụng khác. Vui lòng đóng tệp tin đó lại hoặc chọn nơi lưu khác và thử lại."
        else:
            friendly_msg = f"Lỗi định dạng tài liệu: {err_msg}"
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
        # result có thể là string hoặc None
        return result if result else None


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
