import sys
import os
# Tambahkan path root agar bisa import utils jika folder terpisah
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import base64

import tempfile
import datetime

import requests
import json
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

# Excel
from openpyxl import Workbook

# PDF & Images
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import black, white, gray, Color
from reportlab.pdfbase.pdfmetrics import stringWidth

# Utils (Asumsi Anda punya file utils/db.py dan utils/storage.py)
# Jika tidak, Anda bisa menggabungkan kode koneksi DB/R2 di sini langsung.
from utils.db import get_members_collection
from utils.storage import upload_bytes_to_r2
from utils.helpers import flatten_object # Fungsi flatten yang kita buat sebelumnya

app = Flask(__name__)
CORS(app)

# --- HELPER FUNCTIONS ---

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

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "service": "Himatika Python Worker",
        "status": "Online",
        "timestamp": datetime.datetime.now().isoformat(),
        "endpoints": {
            "excel_export": "/api/sheet/export",
            "excel_import": "/api/sheet/import",
            "pdf_signature": "/api/sign/process",
            "pdf_search_text": "/api/pdf/search-text",
            "activiness_letter": "/api/pdf/activiness-letter",
            "ticket_generator": "/api/pdf/ticket"
        }
    })
# --- ENDPOINTS ---

# 0. EXCEL IMPORT
# 0. EXCEL IMPORT
from openpyxl import load_workbook
import zxingcpp
from PIL import Image

@app.route('/api/pdf/scan-qr', methods=['POST'])
def scan_qr():
    try:
        # Check if file is present
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Save to temp
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            file.save(temp_pdf.name)
            temp_pdf_path = temp_pdf.name

        doc = fitz.open(temp_pdf_path)
        found_data = None
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Improve resolution (3x zoom)
            zoom = 3.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert PyMuPDF Pixmap to PIL Image
            if pix.n >= 4:
                mode = "RGBA"
            elif pix.n == 3:
                mode = "RGB"
            else:
                mode = "L"
            
            img_pil = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

            # Detect attempts with zxing-cpp using Pillow Images
            attempts = [img_pil]
            
            # 2. Grayscale
            if mode != 'L':
                img_gray = img_pil.convert('L')
                attempts.append(img_gray)
            else:
                img_gray = img_pil
            
            # 3. Binary Threshold (Simulation of cv2.threshold)
            # Threshold = 128
            img_binary = img_gray.point(lambda p: 255 if p > 128 else 0, mode='1')
            attempts.append(img_binary)

            # 4. Darker Threshold
            img_binary_dark = img_gray.point(lambda p: 255 if p > 100 else 0, mode='1')
            attempts.append(img_binary_dark)
            
            for img_check in attempts:
                try:
                    results = zxingcpp.read_barcodes(img_check)
                    if results:
                        for result in results:
                            if result.format == zxingcpp.BarcodeFormat.QRCode:
                                found_data = result.text
                                break
                    if found_data:
                        break
                except Exception as e:
                    # Continue if decoding fails for one variant
                    continue
            
            if found_data:
                break
                
        doc.close()
        os.remove(temp_pdf_path)

        if found_data:
             return jsonify({'status': 'found', 'data': found_data})
        else:
             return jsonify({'status': 'not_found'})

    except Exception as e:
        print(f"Error scanning QR: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/sheet/import', methods=['POST'])
def import_generic_sheet():
    try:
        if 'file' not in request.files:
             return jsonify({"error": "No file part"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No selected file"}), 400

        wb = load_workbook(file)
        ws = wb.active # Default to first sheet or active one
        
        data = []
        headers = []
        
        # Read headers (Row 1)
        for cell in ws[1]:
            headers.append(cell.value)
            
        # Read data (Row 2 onwards)
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_data = {}
            for i, cell_value in enumerate(row):
                if i < len(headers):
                    header = headers[i]
                    if header is None: continue

                    # Handle nested keys (e.g. "user.name")
                    keys = str(header).split('.')
                    current_ref = row_data
                    
                    for k_idx, key in enumerate(keys):
                        if k_idx == len(keys) - 1:
                            # Last key, assign value
                            current_ref[key] = cell_value
                        else:
                            # Create nested dict if not exists
                            if key not in current_ref:
                                current_ref[key] = {}
                            current_ref = current_ref[key]
                            
            data.append(row_data)
            
        return jsonify({
            "success": True,
            "data": data
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# 0.5 PDF TEXT SEARCH (For Auto Signature)
@app.route('/api/pdf/search-text', methods=['POST'])
def search_text_pdf():
    try:
        body = request.json
        pdf_url = body.get('pdf')
        search_text = body.get('text')
        
        if not pdf_url or not search_text:
             return jsonify({"error": "Missing pdf url or search text"}), 400
             
        # Download PDF
        response = requests.get(pdf_url)
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch PDF"}), 400
            
        pdf_bytes = response.content
        
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        matches = []
        
        for page_num, page in enumerate(doc):
            # search for text
            text_instances = page.search_for(search_text)
            
            for inst in text_instances:
                # inst is Rect(x0, y0, x1, y1)
                # Frontend expects: { page, x, y, width, height }
                # NOTE: PyMuPDF coordinates are Top-Left (0,0). 
                # PDF.js also uses Top-Left usually, but sometimes Bottom-Left depending on view.
                # app/pages/signatures/[id].vue logic:
                # const y = viewport.height - transform[5]; // Konversi Y dari bawah ke atas if using PDF.js raw
                # But here we return coordinates. Let's return standard Top-Left X, Y.
                
                matches.append({
                    "page": page_num + 1, # 1-based index for frontend consistency
                    "x": inst.x0,
                    "y": inst.y0, # Top
                    "width": inst.width,
                    "height": inst.height,
                    "text": search_text
                })
                
        doc.close()
        
        return jsonify(matches)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# 1. GENERIC EXCEL EXPORT
@app.route('/api/sheet/export', methods=['POST'])
def export_generic_sheet():
    try:
        body = request.json
        complex_data = body.get('data', [])
        title = body.get('title', 'Export')
        headers_input = body.get('headers', None)

        wb = Workbook()
        ws = wb.active
        ws.title = title[:30] # Excel sheet limit

        if not complex_data:
            return jsonify({"message": "No data"}), 400

        if headers_input:
            headers = headers_input
        else:
            all_headers = set()
            for item in complex_data:
                flat = flatten_object(item)
                all_headers.update(flat.keys())
            headers = list(all_headers)
            
        ws.append(headers)
        
        for item in complex_data:
            flat = flatten_object(item)
            row = [flat.get(h, "") for h in headers]
            ws.append(row)

        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)

        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{title}.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 2. PDF SIGNATURE PROCESSOR (Auto-Detect / Manual Location)
@app.route('/api/sign/process', methods=['POST'])
def process_sign_overlay():
    try:
        body = request.json
        pdf_url = body.get('pdf')
        output_path = body.get('outputBlobPath')
        qr_value = body.get('qrValue')
        
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

        # Logika Penentuan Lokasi
        target_locations = manual_locations

        if not target_locations:
            return jsonify({"error": "No signature location found."}), 400

        # Proses Overlay (PyPDF + ReportLab)
        reader = PdfReader(input_pdf_stream)
        writer = PdfWriter()

        # Generate QR Code
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, border=1, box_size=10)
        qr.add_data(qr_value)
        qr.make(fit=True)
        qr_pil = qr.make_image(fill_color="black", back_color="transparent")
        
        qr_bytes = io.BytesIO()
        qr_pil.save(qr_bytes, format='PNG')
        qr_bytes.seek(0)
        qr_img = ImageReader(qr_bytes)

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

                    # Convert Top-Left (Frontend/Fitz) ke Bottom-Left (ReportLab)
                    y_bot = p_h - (y_top + h) + 8 
                    
                    can.drawImage(qr_img, x, y_bot, width=w, height=h, mask='auto')

                can.save()
                packet.seek(0)
                page.merge_page(PdfReader(packet).pages[0])

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

# 3. GENERATE ACTIVINESS LETTER (Dengan Return Lokasi Tanda Tangan)
@app.route('/api/pdf/activiness-letter', methods=['POST'])
def generate_activiness_letter():
    try:
        body = request.json
        member_data = body.get('member')
        point_data = body.get('point')
        chairman_data = body.get('chairman')
        secretary_data = body.get('secretary')
        doc_number = body.get('docNumber')
        period = body.get('period')
        config_data = body.get('config', {})
        
        if not all([member_data, point_data, chairman_data, secretary_data, doc_number, period]):
            return jsonify({"error": "Incomplete data"}), 400

        # Setup PDF
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        margin = 40
        def draw_header(c, y_pos):
            # Base Dir untuk Assets
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            def draw_logo_boxed(path, x, y, max_w, max_h, align='left'):
                try:
                    if not os.path.exists(path):
                        return
                    img = ImageReader(path)
                    iw, ih = img.getSize()
                    aspect = iw / float(ih)
                    
                    # Calculate new dimensions fitting in box
                    if aspect > 1: # Wide
                        w = min(iw, max_w)
                        h = w / aspect
                        if h > max_h:
                            h = max_h
                            w = h * aspect
                    else: # Tall or Square
                        h = min(ih, max_h)
                        w = h * aspect
                        if w > max_w:
                            w = max_w
                            h = w / aspect
                    
                    # Draw
                    draw_x = x
                    if align == 'right':
                        draw_x = x + (max_w - w) # Align right within the box (or just x if x is right edge... logic below)
                    elif align == 'center':
                        draw_x = x + (max_w - w)/2
                    
                    # Note: x passed to this function for 'right' logo is usually the left edge of the right area
                    # But in previous code: width - margin - 60. So that is the LEFT edge of the 60px box.
                    
                    c.drawImage(img, draw_x, y + (max_h - h)/2, width=w, height=h, mask='auto')
                except Exception as e:
                    print(f"Error drawing logo {path}: {e}")

            # Logo Left (ITSNU)
            logo_itsnu_path = os.path.join(base_dir, 'assets', 'img', 'itsnu-logo.png')
            draw_logo_boxed(logo_itsnu_path, margin, y_pos - 15, 80, 80, align='left')

            # Logo Right (HIMATIKA)
            logo_hima_path = os.path.join(base_dir, 'assets', 'img', 'logo.png')
            # Right logo position: width - margin - 80 is the LEFT x of the right logo box
            draw_logo_boxed(logo_hima_path, width - margin - 80, y_pos - 15, 80, 80, align='right')

            # Center Text
            text_y = y_pos + 50
            c.setFont("Times-Bold", 12)
            c.drawCentredString(width/2, text_y, config_data.get('name', 'Himatika'))
            c.setFont("Times-Bold", 12)
            c.drawCentredString(width/2, text_y - 12, "FAKULTAS SAINS DAN TEKNOLOGI")
            c.drawCentredString(width/2, text_y - 24, "INSTITUT TEKNOLOGI DAN SAINS NU PEKALONGAN")
            c.setFont("Times-Bold", 12)
            c.drawCentredString(width/2, text_y - 36, period)
            
            # Address & Contact
            c.setFont("Times-Italic", 9)
            address_str = config_data.get('address', "Jl. Kampus No. 1, Gedung PKM Lt. 2")
            phone_str = config_data.get('phone', '+62 812-3456-7890')
            email_str = config_data.get('email', 'himatika@unvc.ac.id')
            
            contact_line1 = f"Sekretariat: {address_str}"
            contact_line2 = f"Narahubung: {phone_str} | Surel: {email_str}"

            c.drawCentredString(width/2, text_y - 45, contact_line1)
            c.drawCentredString(width/2, text_y - 56, contact_line2)

            # Double Line
            line_y = y_pos - 15
            c.setLineWidth(2)
            c.line(margin, line_y, width - margin, line_y)
            c.setLineWidth(0.5)
            c.line(margin, line_y - 2, width - margin, line_y - 2)
            
            return line_y - 20

        # --- PAGE 1: SURAT ---
        header_bottom = draw_header(c, height - 100)

        # Judul
        title_y = header_bottom
        c.setFont("Times-Bold", 12)
        c.drawCentredString(width/2, title_y, "Surat Keterangan Aktif")
        
        c.setFont("Times-Bold", 12) 
        # Manual underline
        org_text = "Himpunan Mahasiswa Informatika"
        text_w = stringWidth(org_text, "Times-Bold", 12)
        c.drawCentredString(width/2, title_y - 18, org_text)
        c.setLineWidth(1)
        c.line((width/2) - (text_w/2), title_y - 20, (width/2) + (text_w/2), title_y - 20)
        
        c.setFont("Times-Roman", 12)
        c.drawCentredString(width/2, title_y - 35, doc_number)

        # Body
        body_y = title_y - 70
        c.setFont("Times-Roman", 12)
        c.drawString(margin, body_y, "Yang bertanda tangan di bawah ini :")
        
        def draw_kv(lbl, val, y):
            c.drawString(margin, y, lbl)
            c.drawString(margin + 100, y, ": " + str(val))
            return y - 18

        body_y -= 25
        body_y = draw_kv("Nama", chairman_data.get('fullName', '-'), body_y)
        body_y = draw_kv("NIM", chairman_data.get('NIM', '-'), body_y)
        body_y = draw_kv("Jabatan", "Ketua Umum", body_y)

        body_y -= 15
        c.drawString(margin, body_y, "Menyatakan dengan sesungguhnya bahwa :")
        body_y -= 25
        
        body_y = draw_kv("Nama", member_data.get('fullName', '-'), body_y)
        body_y = draw_kv("NIM", member_data.get('NIM', '-'), body_y)
        member_class = member_data.get('class', '-')
        body_y = draw_kv("Kelas", member_class, body_y)
        body_y = draw_kv("Semester", point_data.get('semester', '-'), body_y)

        body_y -= 20
        text = f"Adalah mahasiswa yang benar - benar aktif dalam Himpunan Mahasiswa Informatika (HIMATIKA) ITSNU Pekalongan periode {period}."
        body_y = draw_wrapped_text(c, text, margin, body_y, width - 2*margin, "Times-Roman", 12)

        body_y -= 15
        text2 = "Demikian surat keterangan keaktifan mahasiswa ini dibuat sebagaimana mestinya."
        body_y = draw_wrapped_text(c, text2, margin, body_y, width - 2*margin, "Times-Roman", 12)

        if isinstance(period, str) and len(period) == 4 and period.isdigit():
             # If period is just year, try to parse it or just use it? 
             # Actually `period` is string like "2023/2024". Just verify context.
             pass

        footer_y = body_y - 30
        date_obj = datetime.datetime.now()
        date_str = format_date_indo(date_obj)
        c.drawRightString(width - margin, footer_y, f"Pekalongan, {date_str}")
        
        footer_y -= 20
        c.setFont("Times-Bold", 12)
        c.drawCentredString(width/2, footer_y, "HIMPUNAN MAHASISWA INFORMATIKA")
        c.drawCentredString(width/2, footer_y - 12, "INSTITUT TEKNOLOGI DAN SAINS NAHDLATUL ULAMA")
        c.drawCentredString(width/2, footer_y - 24, "PEKALONGAN")

        footer_y -= 50
        c.setFont("Times-Bold", 12)
        c.drawCentredString(width/2, footer_y, "Mengetahui")
        
        sig_y_start = footer_y - 30
        left_x = width / 4
        right_x = (width * 3) / 4
        
        # Jabatan
        c.setFont("Times-Bold", 12)
        c.drawCentredString(left_x, sig_y_start, "Ketua Umum")
        c.drawCentredString(right_x, sig_y_start, "Sekretaris Umum")
        
        # Signature Placeholders (Grey)
        sig_space_y = sig_y_start - 40
        c.setFillColor(Color(0.7, 0.7, 0.7))
        c.setFont("Courier", 9)
        c.setFillColor(black)

        # Nama
        name_y = sig_y_start - 80
        c.setFont("Times-Bold", 12)
        c.drawCentredString(left_x, name_y, chairman_data.get('fullName', ''))
        c.drawCentredString(right_x, name_y, secretary_data.get('fullName', ''))
        
        # Underline Nama
        c.setLineWidth(0.5)
        # Left Name Underline
        nw_l = stringWidth(chairman_data.get('fullName', ''), "Times-Bold", 12)
        c.line(left_x - (nw_l/2), name_y - 2, left_x + (nw_l/2), name_y - 2)
        # Right Name Underline
        nw_r = stringWidth(secretary_data.get('fullName', ''), "Times-Bold", 12)
        c.line(right_x - (nw_r/2), name_y - 2, right_x + (nw_r/2), name_y - 2)

        c.setFont("Times-Roman", 12)
        c.drawCentredString(left_x, name_y - 14, str(chairman_data.get('NIM', '')))
        c.drawCentredString(right_x, name_y - 14, str(secretary_data.get('NIM', '')))

        # Footer Note
        note_y = 40
        c.setFont("Times-Italic", 7)
        c.drawString(margin, note_y, f"*Surat ini dibuat dengan menggunakan sistem Informasi Himpunan Mahasiswa Informatika (HIMATIKA) ITSNU Pekalongan")
        c.drawString(margin, note_y - 8, f"*Untuk verifikasi keaslian surat ini, silakan kunjungi: {os.getenv('public_url')}/verify/scan")
        c.drawString(margin, note_y - 16, f"dan ditandatangani secara elektronik. Surat ini sah dan berlaku sebagai bukti keaktifan mahasiswa dalam organisasi.")

        # --- SIGNATURE CALCULATIONS ---
        qr_size = 60
        # Signature is roughly centered on sig_space_y
        qr_bottom_y = sig_space_y - (qr_size / 2)
        fe_y = height - (qr_bottom_y + qr_size) # Top-Left System

        sig_locations = [
            {
                "page": 1,
                "role": "Chairman",
                "nim": str(chairman_data.get('NIM')),
                "x": left_x - (qr_size/2),
                "y": fe_y,
                "width": qr_size,
                "height": qr_size
            },
            {
                "page": 1,
                "role": "Secretary",
                "nim": str(secretary_data.get('NIM')),
                "x": right_x - (qr_size/2),
                "y": fe_y,
                "width": qr_size,
                "height": qr_size
            }
        ]
        
        c.showPage()

        # --- PAGE 2: LAMPIRAN (Table) ---
        header_bottom = draw_header(c, height - 100)
        
        curr_y = header_bottom - 20
        c.setFont("Times-Italic", 12)
        c.drawString(margin, curr_y, "Lampiran")
        
        curr_y -= 25
        c.setFont("Times-Roman", 12)
        c.drawString(margin, curr_y, "Daftar keaktifan mahasiswa :")
        
        # Table Header
        curr_y -= 20
        table_y = curr_y
        c.setFont("Times-Bold", 12)
        c.drawString(margin, table_y, "Kategori")
        c.drawString(width/2, table_y, "Jumlah")
        
        c.setLineWidth(1)
        c.line(margin, table_y - 4, width - margin, table_y - 4)
        c.line(margin, table_y + 12, width - margin, table_y + 12) # Upper Line

        # Table Rows
        activities = point_data.get('activities', {})
        agendas = activities.get('agendas', {})
        
        rows = [
            ("Panitia Agenda", agendas.get('committees', 0)),
            ("Peserta Agenda", agendas.get('participants', 0)),
            ("Proyek", activities.get('projects', 0)),
            ("Aspirasi", activities.get('aspirations', 0))
        ]
        
        row_y = table_y - 20
        c.setFont("Times-Roman", 12)
        
        for name, count in rows:
            c.drawString(margin, row_y, name)
            c.drawString(width/2, row_y, str(count))
            row_y -= 18

        c.showPage()
        c.save()
        buffer.seek(0)

        # Upload
        filename = f"{member_data.get('NIM')}/SKA_{point_data.get('semester')}.pdf"
        r2_key = f"documents/activiness-letter/{filename}"
        public_url = upload_bytes_to_r2(buffer.getvalue(), "application/pdf", r2_key)

        return jsonify({
            "success": True,
            "url": public_url,
            "filename": filename,
            "signatureLocations": sig_locations # RETURN INI KE FRONTEND
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

def rupiah_format(angka, with_prefix=False, desimal=2):
    # Format number with comma as thousands separator and dot as decimal
    s = "{:,.{}f}".format(angka, desimal)
    # Swap separators: 1,234.56 -> 1.234,56
    rupiah = s.replace(',', 'v').replace('.', ',').replace('v', '.')
    if with_prefix:
        return "Rp. {}".format(rupiah)
    return rupiah
    
# 4. GENERATE TICKET
@app.route('/api/pdf/ticket', methods=['POST'])
def generate_ticket():
    try:
        body = request.json
        agenda = body.get('agenda')
        amount = body.get('amount')
        participant = body.get('participant')
        role = body.get('role', 'participant')

        # Ukuran Ticket (Landscape)
        width, height = 600, 250
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=(width, height))
        
        # --- CONFIG ---
        primary_color = Color(0, 141/255, 211/255) # Blue #008DD3
        secondary_color = Color(252/255, 132/255, 17/255) # Orange #FC8411
        text_color = Color(0.2, 0.2, 0.2)
        stub_x = 450
        
        # --- BACKGROUND & SHAPES ---
        # 1. Left Accent (Rounded)
        c.setFillColor(primary_color)
        p = c.beginPath()
        p.moveTo(0, 0)
        p.lineTo(40, 0)
        p.lineTo(40, height)
        p.lineTo(0, height)
        c.drawPath(p, fill=1, stroke=0)
        
        # Semi-circle cutout left
        c.setFillColor(white)
        c.circle(0, height/2, 15, fill=1, stroke=0)

        # 2. Right Stub Background
        c.setFillColor(primary_color)
        c.rect(stub_x, 0, width - stub_x, height, fill=1, stroke=0)
        
        # Semi-circle cutout divider
        c.setFillColor(white)
        c.circle(stub_x, height, 10, fill=1, stroke=0) # Top
        c.circle(stub_x, 0, 10, fill=1, stroke=0) # Bottom

        # dashed line
        c.setStrokeColor(white)
        c.setLineWidth(2)
        c.setDash(4, 4)
        c.line(stub_x, 20, stub_x, height - 20)
        c.setDash([]) # Reset

        # --- CONTENT: LEFT BAND ---
        c.saveState()
        c.translate(25, height/2)
        c.rotate(90)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(0, 0, "NO. 001") # Placeholder Number
        c.restoreState()

        # --- CONTENT: MAIN AREA ---
        c.setFillColor(text_color)
        
        # Title
        c.setFont("Helvetica-Bold", 24)
        title = agenda.get('title', 'AGENDA').upper()
        # Wrap title if long? 
        c.drawString(60, height - 50, title)
        
        c.setFont("Helvetica", 10)
        c.setFillColor(secondary_color)
        c.drawString(60, height - 65, "HIMATIKA EVENT TICKET")
        
        # Main Info
        c.setFillColor(text_color)
        c.setFont("Helvetica-Bold", 14)
        member_name = participant.get('member', {}).get('fullName', 'Peserta').upper()
        c.drawString(60, height - 100, member_name)
        
        c.setFont("Helvetica", 10)
        c.setFillColor(gray)
        c.drawString(60, height - 115, role.upper())

        # Date Label
        c.setFillColor(secondary_color)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(stub_x - 20, height - 50, "DATE")
        c.setFillColor(text_color)
        c.setFont("Helvetica", 10)
        # Parse Date start - end
        date_str = agenda.get('date', {}).get('start', '2025-01-01')[:10] + " - " + agenda.get('date', {}).get('end', '2025-01-01')[:10]
        c.drawRightString(stub_x - 20, height - 65, date_str)

        # QR Code (Middle)
        qr_payload = {"id": participant.get('_id'), "role": role}
        qr = qrcode.make(json.dumps(qr_payload))
        qr_mem = io.BytesIO()
        qr.save(qr_mem, format='PNG')
        qr_mem.seek(0)
        qr_img = ImageReader(qr_mem)
        
        c.drawImage(qr_img, 300, 70, 90, 90)

        # Bottom Boxes (Price/Type)
        def draw_box(x, y, text, bg_color):
            c.setFillColor(bg_color)
            c.rect(x, y, 100, 30, fill=1, stroke=0)
            c.setFillColor(white if bg_color == primary_color else text_color)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(x + 50, y + 10, text)

        draw_box(25, 60, role.upper(), primary_color)
        draw_box(135, 60, rupiah_format(amount, True), secondary_color)



        # --- CONTENT: RIGHT STUB ---
        c.saveState()
        c.translate(stub_x + 10, height/2)
        
        c.setFillColor(text_color)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(0, 40, "ADMIT ONE")
        
        c.setFont("Helvetica", 10)
        c.drawString(0, 10, member_name[:15] + "...")
        
        # Small QR for stub
        qr = qrcode.QRCode(box_size=2, border=0)
        qr.add_data(body.get('id', 'TICKET'))
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        
        qr_bytes = io.BytesIO()
        img_qr.save(qr_bytes, format='PNG')
        qr_bytes.seek(0)
        
        c.drawImage(ImageReader(qr_bytes), 0, -60, 50, 50)
        
        c.restoreState()

        c.save()
        buffer.seek(0)
        
        # Return PDF directly
        buffer.seek(0)
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"{body.get('id', 'ticket')}.pdf"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# 5. TOOLS
@app.route('/api/tools/qr', methods=['POST'])
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

@app.route('/api/tools/compress-image', methods=['POST'])
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



if __name__ == '__main__':
    app.run(port=5000, debug=True)