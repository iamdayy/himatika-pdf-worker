import sys
import os
import io
import base64
import tempfile
import datetime
import requests
import json
import fitz
from pypdf import PdfReader, PdfWriter
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import black, white, gray, Color
from reportlab.pdfbase.pdfmetrics import stringWidth
from openpyxl import Workbook
from PIL import Image
import threading
import subprocess
from flask import Blueprint, request, send_file, jsonify
from utils.db import get_members_collection
from utils.storage import upload_bytes_to_r2, get_s3_client
from utils.helpers import flatten_object

# We will need the helper functions here or in utils/helpers.py
# For now, let's just dump them if needed, or better, we moved them to utils/helpers.py in another branch, but we are on main branch.
# Let's extract the helpers from index.py (lines 46 - 165)

bp = Blueprint('tools', __name__)

def month_to_roman(date_obj):
    month = date_obj.month
    roman_numerals = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
    return roman_numerals[month - 1]

def format_date_indo(date_obj):
    months = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    day = date_obj.day
    month = months[date_obj.month - 1]
    year = date_obj.year
    return f"{day} {month} {year}"

def draw_wrapped_text(canvas_obj, text, x, y, max_width, font_name, font_size, line_height=None):
    """
    Menggambar teks yang otomatis turun baris jika melebihi max_width.
    """
    if line_height is None:
        line_height = font_size + 4

    words = text.split(" ")
    current_line = []
    current_y = y
    
    canvas_obj.setFont(font_name, font_size)

    for word in words:
        test_line = " ".join(current_line + [word])
        width = stringWidth(test_line, font_name, font_size)
        
        if width <= max_width:
            current_line.append(word)
        else:
            canvas_obj.drawString(x, current_y, " ".join(current_line))
            current_line = [word]
            current_y -= line_height

    if current_line:
        canvas_obj.drawString(x, current_y, " ".join(current_line))
    
    return current_y


def draw_signature_box(c, x, y_top, w, h, p_h,
                        sig_name='', sig_as='', sig_img=None, overlap=False):
    """
    Layout kotak tanda tangan:
        ┌─────────────────┐
        │  Jabatan (atas) │  ← 15% h, font min 12pt
        ├─────────────────┤
        │  [TTD / QR]     │  ← gambar besar, ujung bawahnya
        │    ↕ overlap    │    menumpangi area nama (~40%)
        ├─────────────────┤  ← garis tipis
        │  Nama (bawah)   │  ← 20% h, font min 12pt
        └─────────────────┘
    Gambar digambar TERAKHIR sehingga secara visual
    "tinta tanda tangan" berada di atas teks nama.
    sig_img : ImageReader-compatible, atau None
    """
    role_h   = h * 0.15   # jabatan di atas
    name_h   = h * 0.20   # nama di bawah
    img_h    = h - role_h - name_h   # sisa → gambar

    # Koordinat ReportLab (bottom-left origin)
    y_bot    = p_h - (y_top + h)
    name_y0  = y_bot               # bawah area nama
    name_y1  = y_bot + name_h      # batas atas nama / batas bawah garis
    img_y0   = name_y1             # bawah area gambar (tanpa overlap)
    img_y1   = img_y0 + img_h      # atas area gambar / batas bawah jabatan
    role_y0  = img_y1

    c.saveState()

    # ── 1. Jabatan teks (atas) ─────────────────────────────────
    if sig_as:
        role_fs = max(12, min(14, role_h * 0.60))
        c.setFont('Helvetica', role_fs)
        c.setFillColor(Color(0.2, 0.2, 0.2))
        role_text_y = role_y0 + (role_h - role_fs) / 2
        c.drawCentredString(x + w / 2, role_text_y, sig_as[:50])

    # ── 2. Garis atas (antara jabatan & gambar) ────────────────
    c.setStrokeColor(Color(0.5, 0.5, 0.5))
    c.setLineWidth(0.3)
    c.line(x, img_y1, x + w, img_y1)

    # ── 3. Garis bawah (antara gambar & nama) ─────────────────
    c.setStrokeColor(Color(0.5, 0.5, 0.5))
    c.setLineWidth(0.5)
    c.line(x, name_y1, x + w, name_y1)

    # ── 4. Nama teks (bawah, di bawah garis) ──────────────────
    if sig_name:
        name_fs = max(12, min(14, name_h * 0.60))
        c.setFont('Helvetica-Bold', name_fs)
        c.setFillColor(Color(0.1, 0.1, 0.1))
        name_text_y = name_y0 + (name_h - name_fs) / 2
        c.drawCentredString(x + w / 2, name_text_y, sig_name[:40])

    # ── 5. Gambar TTD / QR (digambar TERAKHIR → di atas nama) ──
    if sig_img is not None:
        if overlap:
            # Wet signature: gambar tumpang-tindih ke area nama (50% name_h)
            overlap_px  = name_h * 0.50
            draw_y      = img_y0 - overlap_px
            draw_h      = img_h + overlap_px
        else:
            # QR: tetap di dalam area gambar, tidak tumpang tindih
            draw_y  = img_y0
            draw_h  = img_h
        side     = min(w * 0.92, draw_h)
        draw_x   = x + (w - side) / 2
        center_y = draw_y + (draw_h - side) / 2
        c.drawImage(sig_img, draw_x, center_y,
                    width=side, height=side,
                    mask='auto', preserveAspectRatio=False)

    c.restoreState()

# 5. TOOLS
@bp.route('/api/tools/qr', methods=['POST'])
def generate_qr_tool():
    try:
        body = request.json
        text = body.get('text')
        
        if not text:
            return jsonify({"error": "No text provided"}), 400

        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        qr.add_data(text)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
        
        return jsonify({
            "success": True,
            "dataUrl": f"data:image/png;base64,{img_str}"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route('/api/tools/compress-image', methods=['POST'])
def compress_image_tool():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        quality = int(request.form.get('quality', 80))
        max_width = request.form.get('maxWidth')
        max_height = request.form.get('maxHeight')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Open image using Pillow
        img = Image.open(file.stream)
        
        # Convert to RGB if necessary (e.g. for JPEG saving)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        # Resize if needed
        if max_width or max_height:
            w, h = img.size
            ratio = min(
                int(max_width)/w if max_width else 1, 
                int(max_height)/h if max_height else 1
            )
            # Only downscale
            if ratio < 1:
                new_size = (int(w*ratio), int(h*ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                
        # Compress
        output_buffer = io.BytesIO()
        # Default to JPEG for compression efficiency unless original was transparency heavy, 
        # but here we force JPEG/WebP as requested or just standard JPEG for simplicity 
        # based on 'browser-image-compression' replacement context usually implying JPEG/WebP.
        # Let's return JPEG for now as it's most common for 'compression'.
        
        img.save(output_buffer, format='JPEG', quality=quality, optimize=True)
        output_buffer.seek(0)
        
        # Encode to base64 to return easily (or binary if preferred, but base64 is easier for standard JSON API)
        # However, `customReadMultipartFormData` in Nuxt expects a buffer. 
        # Sending base64 is safer for JSON response, or we can return raw bytes with correct mimetype.
        # Let's return base64 and decode in Nuxt to Buffer.
        
        img_str = base64.b64encode(output_buffer.getvalue()).decode("utf-8")
        
        return jsonify({
            "success": True,
            "data": img_str,
            "mime": "image/jpeg"
        })

    except Exception as e:
        print(f"Error compressing image: {e}")
        return jsonify({'error': str(e)}), 500


@bp.route('/api/tools/upload-image', methods=['POST'])
def upload_image_to_r2():
    """Upload an image to R2 and return its public URL.
    Preserves PNG transparency — important for wet signature images.
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        quality = int(request.form.get('quality', 95))
        img = Image.open(file.stream)
        out = io.BytesIO()

        # Preserve transparency for PNG signatures
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            img.convert('RGBA').save(out, format='PNG', optimize=True)
            content_type = 'image/png'
            ext = 'png'
        else:
            img.convert('RGB').save(out, format='JPEG', quality=quality, optimize=True)
            content_type = 'image/jpeg'
            ext = 'jpeg'

        out.seek(0)
        import uuid
        key = f"uploads/signatures/{uuid.uuid4().hex}.{ext}"
        public_url = upload_bytes_to_r2(out.getvalue(), content_type, key)
        return jsonify({'success': True, 'url': public_url})

    except Exception as e:
        print(f"upload_image_to_r2 error: {e}")
        return jsonify({'error': str(e)}), 500

