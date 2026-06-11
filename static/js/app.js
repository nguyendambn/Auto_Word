/* =====================================================
   Auto-Word — Main Application JavaScript
   ===================================================== */

document.addEventListener('DOMContentLoaded', () => {
    'use strict';

    // ── DOM References ───────────────────────────────
    const uploadZone         = document.getElementById('uploadZone');
    const fileInput          = document.getElementById('fileInput');
    const fileListContainer  = document.getElementById('fileListContainer');
    const themeToggle        = document.getElementById('themeToggle');

    const marginTop          = document.getElementById('marginTop');
    const marginBottom       = document.getElementById('marginBottom');
    const marginLeft         = document.getElementById('marginLeft');
    const marginRight        = document.getElementById('marginRight');
    const valMarginTop       = document.getElementById('val-marginTop');
    const valMarginBottom    = document.getElementById('val-marginBottom');
    const valMarginLeft      = document.getElementById('val-marginLeft');
    const valMarginRight     = document.getElementById('val-marginRight');
    const lblMarginTop       = document.getElementById('lbl-margin-top');
    const lblMarginBottom    = document.getElementById('lbl-margin-bottom');
    const lblMarginLeft      = document.getElementById('lbl-margin-left');
    const lblMarginRight     = document.getElementById('lbl-margin-right');

    const fontName           = document.getElementById('fontName');
    const bodySize           = document.getElementById('bodySize');
    const lineSpacing        = document.getElementById('lineSpacing');
    const spaceBefore        = document.getElementById('spaceBefore');
    const spaceAfter         = document.getElementById('spaceAfter');
    const alignment          = document.getElementById('alignment');
    const firstLineIndent    = document.getElementById('firstLineIndent');
    const contextualSpacing = document.getElementById('contextualSpacing');

    const autoNumberHeadings = document.getElementById('autoNumberHeadings');
    const heading1Size       = document.getElementById('heading1Size');
    const heading1Bold       = document.getElementById('heading1Bold');
    const heading1Uppercase  = document.getElementById('heading1Uppercase');
    const heading2Size       = document.getElementById('heading2Size');
    const heading2Bold       = document.getElementById('heading2Bold');
    const heading3Size       = document.getElementById('heading3Size');
    const heading3Italic     = document.getElementById('heading3Italic');

    const addPageNumbers     = document.getElementById('addPageNumbers');
    const formatCover        = document.getElementById('formatCover');

    const formatBtn          = document.getElementById('formatBtn');
    const loadingOverlay     = document.getElementById('loadingOverlay');
    const progressStep       = document.getElementById('progressStep');
    const progressBarInner   = document.getElementById('progressBarInner');

    const toast              = document.getElementById('toast');
    const toastTitle         = document.getElementById('toastTitle');
    const toastMsg           = document.getElementById('toastMsg');
    const toastIcon          = document.getElementById('toastIcon');

    const presetBtns         = document.querySelectorAll('.preset-btn');
    const tabBtns            = document.querySelectorAll('.tab-btn');
    const tabContents        = document.querySelectorAll('.tab-content');

    // ── Constants ────────────────────────────────────
    const PX_SCALE = 2.28; // mm → px conversion for A4 preview

    // ── State ────────────────────────────────────────
    let selectedFiles = [];
    let toastTimer   = null;

    // ── Presets ──────────────────────────────────────
    const PRESETS = {
        dakltn: {
            fontName: 'Times New Roman', bodySize: 14, lineSpacing: 1.5,
            spaceBefore: 0, spaceAfter: 0,
            marginTop: 25, marginBottom: 20, marginLeft: 35, marginRight: 20,
            firstLineIndent: 10, autoNumberHeadings: true,
            heading1Size: 14, heading1Bold: true, heading1Uppercase: true,
            heading2Size: 14, heading2Bold: true,
            heading3Size: 14, heading3Italic: false,
            formatAdminParts: false, addPageNumbers: true,
            contextualSpacing: true, formatCover: false
        }
    };

    // ── Helpers ──────────────────────────────────────
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1048576).toFixed(2) + ' MB';
    }

    // ── Toast ────────────────────────────────────────
    function showToast(title, message, type) {
        toastTitle.textContent = title;
        toastMsg.textContent   = message;

        toast.className = 'toast ' + type;
        toastIcon.className = type === 'success'
            ? 'toast-icon fa-solid fa-circle-check'
            : 'toast-icon fa-solid fa-circle-exclamation';

        toast.style.display = 'flex';

        if (toastTimer) clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toast.style.display = 'none';
        }, 4000);
    }

    // ── Theme Toggle ─────────────────────────────────
    const htmlEl = document.documentElement;
    const savedTheme = localStorage.getItem('app-theme') || 'dark';
    htmlEl.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    themeToggle.addEventListener('click', () => {
        const currentTheme = htmlEl.getAttribute('data-theme');
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        htmlEl.setAttribute('data-theme', newTheme);
        localStorage.setItem('app-theme', newTheme);
        updateThemeIcon(newTheme);
    });

    function updateThemeIcon(theme) {
        themeToggle.innerHTML = theme === 'dark' 
            ? '<i class="fa-solid fa-sun"></i>' 
            : '<i class="fa-solid fa-moon"></i>';
    }

    // ══════════════════════════════════════════════════
    //  FILE HANDLING
    // ══════════════════════════════════════════════════

    uploadZone.addEventListener('click', () => fileInput.click());

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length) handleFiles(files);
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) handleFiles(fileInput.files);
    });

    function handleFiles(files) {
        let hasInvalidFile = false;
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (!file.name.toLowerCase().endsWith('.docx')) {
                hasInvalidFile = true;
                continue;
            }
            if (!selectedFiles.some(f => f.name === file.name && f.size === file.size)) {
                selectedFiles.push(file);
            }
        }
        if (hasInvalidFile) {
            showToast('Lỗi định dạng', 'Chỉ chấp nhận tệp .docx', 'error');
        }
        renderFileList();
    }

    function renderFileList() {
        fileListContainer.innerHTML = '';
        if (selectedFiles.length === 0) {
            fileListContainer.style.display = 'none';
            uploadZone.style.display = 'flex';
            formatBtn.disabled = true;
            return;
        }

        // Thêm nút "Xoá tất cả" nếu chọn từ 2 file trở lên
        if (selectedFiles.length > 1) {
            const headerDiv = document.createElement('div');
            headerDiv.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; padding: 0 4px;';
            headerDiv.innerHTML = `
                <span style="font-size: 13px; font-weight: 600; color: var(--text-strong);">Danh sách tệp (${selectedFiles.length})</span>
                <button class="icon-btn remove-btn" id="clearAllBtn" style="width: auto; height: auto; padding: 6px 12px; font-size: 11px; font-weight: 500; display: flex; gap: 4px; align-items: center; border-radius: 6px;" title="Xóa tất cả các tệp">
                    <i class="fa-solid fa-trash-can"></i> Xóa tất cả
                </button>
            `;
            headerDiv.querySelector('#clearAllBtn').addEventListener('click', (e) => {
                e.stopPropagation();
                selectedFiles = [];
                fileInput.value = '';
                renderFileList();
            });
            fileListContainer.appendChild(headerDiv);
        }

        selectedFiles.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <div class="file-info-inner">
                    <div class="file-icon-box"><i class="fa-solid fa-file-word"></i></div>
                    <div class="file-text" style="max-width: 320px;">
                        <span class="file-name" title="${file.name}">${file.name}</span>
                        <span class="file-size">${formatFileSize(file.size)}</span>
                    </div>
                </div>
                <button class="icon-btn remove-btn" data-index="${index}" title="Xóa tệp"><i class="fa-solid fa-xmark"></i></button>
            `;
            
            item.querySelector('.remove-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                selectedFiles.splice(index, 1);
                fileInput.value = '';
                renderFileList();
            });
            
            fileListContainer.appendChild(item);
        });

        const addMoreDiv = document.createElement('div');
        addMoreDiv.style.cssText = 'display: flex; justify-content: center; margin-top: 8px; width: 100%;';
        addMoreDiv.innerHTML = `
            <button class="preset-btn" style="flex-direction: row; padding: 8px 16px; border-style: dashed; width: 100%; justify-content: center; gap: 8px;">
                <i class="fa-solid fa-plus" style="font-size: 14px;"></i> Thêm tệp Word khác
            </button>
        `;
        addMoreDiv.querySelector('button').addEventListener('click', () => {
            fileInput.click();
        });
        fileListContainer.appendChild(addMoreDiv);

        uploadZone.style.display = 'none';
        fileListContainer.style.display = 'flex';
        formatBtn.disabled = false;
        
        if (selectedFiles.length === 1) {
            formatBtn.querySelector('span').textContent = 'Định dạng tự động';
        } else {
            formatBtn.querySelector('span').textContent = `Định dạng ${selectedFiles.length} file`;
        }
    }

    // ══════════════════════════════════════════════════
    //  TAB SWITCHING
    // ══════════════════════════════════════════════════

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(btn.dataset.tab).classList.add('active');
        });
    });

    // ══════════════════════════════════════════════════
    //  PRESET SWITCHING
    // ══════════════════════════════════════════════════

    presetBtns.forEach(card => {
        card.addEventListener('click', () => {
            const presetKey = card.dataset.preset;
            presetBtns.forEach(c => c.classList.remove('active'));
            card.classList.add('active');

            if (presetKey !== 'custom' && PRESETS[presetKey]) {
                applyPreset(PRESETS[presetKey]);
            }
        });
    });

    function applyPreset(p) {
        fontName.value        = p.fontName;
        bodySize.value        = p.bodySize;
        lineSpacing.value     = p.lineSpacing;
        spaceBefore.value     = p.spaceBefore;
        spaceAfter.value      = p.spaceAfter;
        marginTop.value       = p.marginTop;
        marginBottom.value    = p.marginBottom;
        marginLeft.value      = p.marginLeft;
        marginRight.value     = p.marginRight;
        firstLineIndent.value = p.firstLineIndent;
        contextualSpacing.checked = p.contextualSpacing !== undefined ? p.contextualSpacing : true;

        autoNumberHeadings.checked = p.autoNumberHeadings;
        heading1Size.value         = p.heading1Size;
        heading1Bold.checked       = p.heading1Bold;
        heading1Uppercase.checked  = p.heading1Uppercase;
        heading2Size.value         = p.heading2Size;
        heading2Bold.checked       = p.heading2Bold;
        heading3Size.value         = p.heading3Size;
        heading3Italic.checked     = p.heading3Italic;

        addPageNumbers.checked   = p.addPageNumbers;
        formatCover.checked      = p.formatCover !== undefined ? p.formatCover : false;

        updatePreview();
    }

    // Switch to "custom" preset on any manual input change
    function switchToCustom() {
        presetBtns.forEach(c => c.classList.remove('active'));
        document.querySelector('.preset-btn[data-preset="custom"]').classList.add('active');
    }

    const allInputs = [
        fontName, bodySize, lineSpacing, spaceBefore, spaceAfter,
        alignment, firstLineIndent, contextualSpacing,
        marginTop, marginBottom, marginLeft, marginRight,
        autoNumberHeadings, heading1Size, heading1Bold, heading1Uppercase,
        heading2Size, heading2Bold, heading3Size, heading3Italic,
        addPageNumbers, formatCover
    ];

    allInputs.forEach(el => {
        el.addEventListener('input', () => {
            switchToCustom();
            updatePreview();
        });
        el.addEventListener('change', () => {
            switchToCustom();
            updatePreview();
        });
    });

    // ══════════════════════════════════════════════════
    //  A4 LIVE PREVIEW
    // ══════════════════════════════════════════════════

    function updatePreview() {
        // --- Margin values ---
        const mt = parseInt(marginTop.value);
        const mb = parseInt(marginBottom.value);
        const ml = parseInt(marginLeft.value);
        const mr = parseInt(marginRight.value);

        valMarginTop.textContent    = mt;
        valMarginBottom.textContent = mb;
        valMarginLeft.textContent   = ml;
        valMarginRight.textContent  = mr;

    }

    // ══════════════════════════════════════════════════
    //  FORMAT BUTTON — SAVE DIALOG FLOW
    // ══════════════════════════════════════════════════

    formatBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) return;

        let savePath = null;
        let isMultiple = selectedFiles.length > 1;

        if (window.pywebview && window.pywebview.api) {
            try {
                if (isMultiple) {
                    savePath = await window.pywebview.api.select_save_folder();
                } else {
                    const originalName  = selectedFiles[0].name;
                    const dotIdx        = originalName.lastIndexOf('.');
                    const formattedName = `${originalName.substring(0, dotIdx)}_formatted.docx`;
                    savePath = await window.pywebview.api.select_save_path(formattedName);
                }
                if (!savePath) return; // User cancelled
            } catch (err) {
                console.warn('Save dialog error:', err);
            }
        }

        // Show loading overlay
        loadingOverlay.style.display = 'flex';
        progressBarInner.style.width = '10%';
        progressStep.textContent     = 'Đang chuẩn bị đọc tệp tin...';

        // Animated progress steps
        const steps = [
            { width: '25%', text: 'Đang tải tệp tin lên bộ nhớ tạm...' },
            { width: '45%', text: 'Đang phân tích định dạng paragraphs...' },
            { width: '70%', text: 'Đang áp dụng cỡ chữ và căn lề trang...' },
            { width: '90%', text: 'Đang tạo số trang và định dạng bảng biểu...' }
        ];
        let stepIdx = 0;
        const intervalId = setInterval(() => {
            if (stepIdx < steps.length) {
                progressBarInner.style.width = steps[stepIdx].width;
                progressStep.textContent     = steps[stepIdx].text;
                stepIdx++;
            }
        }, 800);

        // Build FormData with all settings
        const formData = new FormData();
        selectedFiles.forEach(file => {
            formData.append('file', file);
        });
        if (savePath) formData.append('save_path', savePath);

        formData.append('font_name',           fontName.value);
        formData.append('body_size',           bodySize.value);
        formData.append('line_spacing',        lineSpacing.value);
        formData.append('space_before',        spaceBefore.value);
        formData.append('space_after',         spaceAfter.value);
        formData.append('alignment',           alignment.value);
        formData.append('first_line_indent',   firstLineIndent.value);
        formData.append('contextual_spacing',   contextualSpacing.checked);

        formData.append('margin_top',          marginTop.value);
        formData.append('margin_bottom',       marginBottom.value);
        formData.append('margin_left',         marginLeft.value);
        formData.append('margin_right',        marginRight.value);

        formData.append('auto_number_headings', autoNumberHeadings.checked);
        formData.append('heading1_size',       heading1Size.value);
        formData.append('heading1_bold',       heading1Bold.checked);
        formData.append('heading1_uppercase',  heading1Uppercase.checked);
        formData.append('heading2_size',       heading2Size.value);
        formData.append('heading2_bold',       heading2Bold.checked);
        formData.append('heading3_size',       heading3Size.value);
        formData.append('heading3_italic',     heading3Italic.checked);

        formData.append('format_admin_parts',  'false');
        formData.append('add_page_numbers',    addPageNumbers.checked);
        formData.append('format_cover',        formatCover.checked);

        try {
            const response = await fetch('/api/format', {
                method: 'POST',
                body: formData
            });
            clearInterval(intervalId);

            if (!response.ok) {
                const contentType = response.headers.get('content-type') || '';
                if (contentType.includes('application/json')) {
                    const errData = await response.json();
                    throw new Error(errData.error || 'Lỗi không xác định');
                } else {
                    throw new Error(`Lỗi máy chủ (${response.status}). File có thể quá lớn hoặc xử lý quá lâu. Vui lòng thử lại.`);
                }
            }

            progressBarInner.style.width = '100%';
            progressStep.textContent     = 'Hoàn tất!';

            if (savePath) {
                // File was saved directly by Python backend
                setTimeout(() => {
                    loadingOverlay.style.display = 'none';
                    showToast('Lưu thành công', `Đã lưu tại: ${savePath}`, 'success');
                }, 600);
            } else {
                // Fallback: download blob
                const blob = await response.blob();
                const url  = window.URL.createObjectURL(blob);
                const a    = document.createElement('a');
                a.href     = url;
                if (isMultiple) {
                    a.download = "AutoWord_Formatted_Files.zip";
                } else {
                    const originalName  = selectedFiles[0].name;
                    const dotIdx        = originalName.lastIndexOf('.');
                    a.download = `${originalName.substring(0, dotIdx)}_formatted.docx`;
                }
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
                setTimeout(() => {
                    loadingOverlay.style.display = 'none';
                    showToast('Thành công', 'Tải xuống thành công!', 'success');
                }, 600);
            }
        } catch (error) {
            clearInterval(intervalId);
            loadingOverlay.style.display = 'none';
            showToast('Lỗi xử lý', error.message, 'error');
        }
    });

    // ── Download Cover Templates (Pywebview Bridge) ─────────────────
    const coverLinks = document.querySelectorAll('a[download]');
    coverLinks.forEach(link => {
        link.addEventListener('click', async (e) => {
            if (window.pywebview && window.pywebview.api) {
                e.preventDefault();
                const filename = link.getAttribute('href').split('/').pop();
                try {
                    const result = await window.pywebview.api.download_cover(filename);
                    if (result && result.success) {
                        showToast('Tải xuống thành công', `Đã lưu tại: ${result.saved_to}`, 'success');
                    } else if (result && result.error) {
                        showToast('Lỗi tải xuống', result.error, 'error');
                    }
                } catch (err) {
                    showToast('Lỗi kết nối', err.message, 'error');
                }
            }
        });
    });

    applyPreset(PRESETS.dakltn);
});
