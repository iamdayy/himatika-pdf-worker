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
import threading
import subprocess
from flask import Blueprint, request, send_file, jsonify
from utils.db import get_members_collection
from utils.storage import upload_bytes_to_r2, get_s3_client
from utils.helpers import flatten_object

# We will need the helper functions here or in utils/helpers.py
# For now, let's just dump them if needed, or better, we moved them to utils/helpers.py in another branch, but we are on main branch.
# Let's extract the helpers from index.py (lines 46 - 165)

bp = Blueprint('sign', __name__)

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

# 2. PDF SIGNATURE PROCESSOR (Auto-Detect / Manual Location)
@bp.route('/api/sign/process', methods=['POST'])
def process_sign_overlay():
    try:
        body = request.json
        pdf_url = body.get('pdf')
        output_path = body.get('outputBlobPath')
        qr_value = body.get('qrValue')
        signer_name = body.get('signerName', '')  # nama member penandatangan
        signer_as   = body.get('signerAs', '')    # jabatan penandatangan
        
        # Opsi: Manual locations ATAU Search Text
        manual_locations = body.get('locations', [])

        if not all([pdf_url, output_path, qr_value]):
            return jsonify({"error": "Missing parameters"}), 400

        # Download PDF
        response = requests.get(pdf_url)
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch PDF"}), 400
        
        pdf_bytes = response.content
        input_pdf_stream = io.BytesIO(pdf_bytes)

        target_locations = manual_locations
        if not target_locations:
            return jsonify({"error": "No signature location found."}), 400

        # Generate QR Code
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, border=1, box_size=10)
        qr.add_data(qr_value)
        qr.make(fit=True)
        qr_pil = qr.make_image(fill_color="black", back_color="white")
        
        qr_bytes = io.BytesIO()
        qr_pil.save(qr_bytes, format='PNG')
        qr_bytes.seek(0)
        qr_img = ImageReader(qr_bytes)

        reader = PdfReader(input_pdf_stream)
        writer = PdfWriter()

        # Group Locations
        locs_by_page = {}
        for loc in target_locations:
            p = int(loc.get('page', 1)) - 1
            if p not in locs_by_page: locs_by_page[p] = []
            locs_by_page[p].append(loc)

        for i, page in enumerate(reader.pages):
            if i in locs_by_page:
                p_w = float(page.mediabox.width)
                p_h = float(page.mediabox.height)
                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(p_w, p_h))

                for loc in locs_by_page[i]:
                    x = float(loc.get('x', 0))
                    y_top = float(loc.get('y', 0))
                    w = float(loc.get('width', 100))
                    h = float(loc.get('height', 100))

                    # Hanya gambar QR, abaikan teks dan garis
                    y_bot_qr = p_h - (y_top + h)
                    # Use a centered square for the QR code to ensure it's not distorted and centered
                    side = min(w, h)
                    draw_x = x + (w - side) / 2
                    draw_y = y_bot_qr + (h - side) / 2
                    can.drawImage(qr_img, draw_x, draw_y, width=side, height=side, mask='auto', preserveAspectRatio=True)

                can.save()
                packet.seek(0)
                page.merge_page(PdfReader(packet).pages[0])
                packet.close()

            writer.add_page(page)

        output_buffer = io.BytesIO()
        writer.write(output_buffer)
        output_buffer.seek(0)
        
        public_url = upload_bytes_to_r2(output_buffer.getvalue(), "application/pdf", output_path)

        return jsonify({
            "success": True, 
            "data": public_url
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    finally:
        if 'input_pdf_stream' in locals(): input_pdf_stream.close()
        if 'qr_bytes' in locals(): qr_bytes.close()
        if 'output_buffer' in locals(): output_buffer.close()

