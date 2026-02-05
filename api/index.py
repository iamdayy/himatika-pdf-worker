
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import requests
import datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import Workbook
import qrcode
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import fitz

# Import utils
from utils.db import get_members_collection
from utils.storage import upload_bytes_to_r2
from utils.helpers import flatten_object, month_to_roman, draw_wrapped_text

from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black, Color

app = Flask(__name__)
CORS(app)

# --- 1. GENERIC EXCEL EXPORT ---
# Replikasi: server/api/sheet/export.post.ts
@app.route('/api/sheet/export', methods=['POST'])
def export_generic_sheet():
    try:
        body = request.json
        complex_data = body.get('data', [])
        title = body.get('title', 'Export')
        headers_input = body.get('headers', None)

        wb = Workbook()
        ws = wb.active
        ws.title = title

        if not complex_data:
            return jsonify({"message": "No data provided"}), 400

        # Tentukan Headers
        if headers_input:
            headers = headers_input
            ws.append(headers)
            for item in complex_data:
                flat_item = flatten_object(item)
                row = [flat_item.get(h, "") for h in headers]
                ws.append(row)
        else:
            # Auto-detect headers dari semua keys
            all_headers = set()
            for item in complex_data:
                flat_item = flatten_object(item)
                all_headers.update(flat_item.keys())
            
            headers = list(all_headers)
            ws.append(headers)
            
            for item in complex_data:
                flat_item = flatten_object(item)
                row = [flat_item.get(h, "") for h in headers]
                ws.append(row)

        # Generate filename
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        filename = f"{title}-{timestamp}.xlsx"

        # Save ke memory buffer
        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)

        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Error generic export: {e}")
        return jsonify({"statusCode": 500, "statusMessage": str(e)}), 500


# --- 2. MEMBER EXPORT ---
# Replikasi: server/api/member/sheet/export.post.ts
@app.route('/api/member/sheet/export', methods=['POST'])
def export_members():
    try:
        body = request.json or {}
        nims = body.get('data', []) # Array of NIMs

        query = {}
        if nims:
            query['NIM'] = {'$in': nims}

        members_col = get_members_collection()
        # Ambil data spesifik sesuai TypeScript
        cursor = members_col.find(
            query, 
            {"NIM": 1, "fullName": 1, "email": 1, "class": 1, "semester": 1, "enteredYear": 1, "status": 1, "_id": 0}
        )
        
        members = list(cursor)
        count = len(members)

        wb = Workbook()
        ws = wb.active
        ws.title = "Member"

        # Header Columns
        headers = ["NIM", "Nama Lengkap", "Email", "Kelas", "Semester", "Tahun Masuk", "Status"]
        keys = ["NIM", "fullName", "email", "class", "semester", "enteredYear", "status"]
        
        ws.append(headers)

        # Mengisi baris
        for m in members:
            row = [m.get(k, "") for k in keys]
            ws.append(row)

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        filename = f"exported-{count}-{timestamp}.xlsx"

        # Save buffer
        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)

        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Error member export: {e}")
        return jsonify({"statusCode": 500, "statusMessage": str(e)}), 500


# --- 4. GENERATE ACTIVINESS LETTER ---
# Replikasi: makeActivinessLetter di useMakeDocs.ts
@app.route('/api/pdf/activiness-letter', methods=['POST'])
def generate_activiness_letter():
    try:
        body = request.json
        
        # Data Input
        user_data = body.get('user') # Data User Login
        point_data = body.get('point') # Data Point/Semester
        organizer = body.get('organizer') # Data Pengurus
        config_data = body.get('config') # Konfigurasi HIMA
        doc_number = body.get('docNumber') # Nomor Surat Lengkap
        
        if not all([user_data, point_data, organizer, config_data, doc_number]):
            return jsonify({"error": "Incomplete data"}), 400

        # Ambil Ketua & Sekretaris
        daily_mgmt = organizer.get('dailyManagement', [])
        chairman = next((dm['member'] for dm in daily_mgmt if 'Ketua' in dm['position'] or 'Chairman' in dm['position']), None)
        secretary = next((dm['member'] for dm in daily_mgmt if 'Sekretaris' in dm['position'] or 'Secretary' in dm['position']), None)

        if not chairman or not secretary:
            return jsonify({"error": "Chairman or Secretary not found"}), 400

        # Setup PDF
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4 # 595.27, 841.89 points
        margin = 40 # Sesuaikan margin (TS code pakai 20 tapi di reportlab 40 biasanya lebih aman)

        # --- FONT SETUP ---
        # ReportLab punya font bawaan: Times-Roman, Times-Bold, Times-Italic
        
        # --- PAGE 1 ---
        
        # 1. Header Images (Logo)
        # Gunakan URL absolut atau file lokal jika ada. Di sini kita asumsikan URL publik.
        logo_hima_url = "https://your-domain.com/img/logo.png" # Ganti dengan URL Logo Anda
        logo_itsnu_url = "https://your-domain.com/img/itsnu-logo.png"
        
        try:
            # Fetch gambar dulu
            logo_hima = ImageReader(logo_hima_url)
            logo_itsnu = ImageReader(logo_itsnu_url)
            
            c.drawImage(logo_hima, margin, height - 100, width=60, height=60, mask='auto')
            c.drawImage(logo_itsnu, width - margin - 60, height - 100, width=60, height=60, mask='auto')
        except:
            print("Warning: Logo not found, skipping...")

        # 2. Header Text (Centered)
        header_y = height - 50
        c.setFont("Times-Bold", 12)
        c.drawCentredString(width/2, header_y, config_data.get('name', 'Himpunan Mahasiswa Informatika'))
        c.drawCentredString(width/2, header_y - 14, "FAKULTAS SAINS DAN TEKNOLOGI")
        c.drawCentredString(width/2, header_y - 28, "INSTITUT TEKNOLOGI DAN SAINS NAHDLATUL ULAMA PEKALONGAN")
        
        # Periode
        period = organizer.get('period', '2025 - 2026')
        c.drawCentredString(width/2, header_y - 42, period)
        
        c.setFont("Times-Italic", 11)
        address = config_data.get('address', '')
        contact = f"narahubung: {config_data.get('contact', {}).get('phone', '')} surel: {config_data.get('contact', {}).get('email', '')}"
        c.drawCentredString(width/2, header_y - 56, address)
        c.drawCentredString(width/2, header_y - 68, contact)

        # 3. Garis Header
        line_y = header_y - 80
        c.setLineWidth(1)
        c.line(margin, line_y, width - margin, line_y)
        c.setLineWidth(2)
        c.line(margin, line_y - 4, width - margin, line_y - 4)

        # 4. Judul Surat
        title_y = line_y - 40
        c.setFont("Times-Bold", 14)
        c.drawCentredString(width/2, title_y, "Surat Keterangan Aktif")
        c.drawCentredString(width/2, title_y - 18, "Himpunan Mahasiswa Informatika")
        
        # Garis Bawah Judul
        text_w = c.stringWidth("Himpunan Mahasiswa Informatika", "Times-Bold", 14)
        c.setLineWidth(1)
        c.line((width - text_w)/2, title_y - 20, (width + text_w)/2, title_y - 20)

        # Nomor Surat
        c.setFont("Times-Roman", 12)
        c.drawCentredString(width/2, title_y - 35, doc_number)

        # 5. Isi Surat (Body)
        body_y = title_y - 70
        c.drawString(margin + 20, body_y, "Yang bertanda tangan di bawah ini :")
        
        def draw_kv(label, value, y_pos):
            c.drawString(margin + 20, y_pos, label)
            c.drawString(margin + 110, y_pos, ":")
            c.drawString(margin + 120, y_pos, str(value))
            return y_pos - 18

        body_y -= 20
        body_y = draw_kv("Nama", chairman.get('fullName', ''), body_y)
        body_y = draw_kv("NIM", chairman.get('NIM', ''), body_y)
        body_y = draw_kv("Jabatan", "Ketua Umum", body_y)

        body_y -= 15
        c.drawString(margin + 20, body_y, "Menyatakan dengan sesungguhnya bahwa :")
        body_y -= 20
        
        member = user_data.get('member', {})
        body_y = draw_kv("Nama", member.get('fullName', ''), body_y)
        body_y = draw_kv("NIM", member.get('NIM', ''), body_y)
        body_y = draw_kv("Kelas", member.get('class', ''), body_y)
        body_y = draw_kv("Semester", point_data.get('semester', ''), body_y)

        body_y -= 20
        statement = f"Adalah mahasiswa yang benar - benar aktif dalam Himpunan Mahasiswa Informatika (HIMATIKA) ITSNU Pekalongan periode {period}."
        
        # Gunakan helper wrap text
        body_y = draw_wrapped_text(c, statement, margin + 20, body_y, width - 2*(margin+20), "Times-Roman", 12)
        
        body_y -= 20
        closing = "Demikian surat keterangan keaktifan mahasiswa ini dibuat sebagaimana mestinya."
        body_y = draw_wrapped_text(c, closing, margin + 20, body_y, width - 2*(margin+20), "Times-Roman", 12)

        # 6. Footer (Tanda Tangan)
        footer_y = body_y - 40
        date_str = datetime.datetime.now().strftime("%d %B %Y") # Perlu format ID manual jika di server non-ID
        c.drawRightString(width - margin, footer_y, f"Pekalongan, {date_str}")
        
        footer_y -= 20
        c.setFont("Times-Bold", 12)
        c.drawCentredString(width/2, footer_y, "HIMPUNAN MAHASISWA INFORMATIKA")
        c.drawCentredString(width/2, footer_y - 14, "INSTITUT TEKNOLOGI DAN SAINS NAHDLATUL ULAMA")
        c.drawCentredString(width/2, footer_y - 28, "PEKALONGAN")
        
        footer_y -= 50
        c.drawCentredString(width/2, footer_y, "Mengetahui")
        
        # Tanda Tangan Kiri (Ketua) & Kanan (Sekretaris)
        sig_y = footer_y - 40
        left_x = width / 4
        right_x = (width * 3) / 4
        
        c.drawCentredString(left_x, sig_y, "Ketua Umum")
        c.drawCentredString(right_x, sig_y, "Sekretaris Umum")
        
        # Teks Signature (Mocking Digital Signature)
        c.setFont("Times-Roman", 12)
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.drawCentredString(left_x, sig_y - 40, f"/{chairman.get('NIM', '')}signature/")
        c.drawCentredString(right_x, sig_y - 40, f"/{secretary.get('NIM', '')}signature/")
        c.setFillColorRGB(0, 0, 0)
        
        # Nama Tanda Tangan
        c.setFont("Times-Bold", 12)
        c.drawCentredString(left_x, sig_y - 70, chairman.get('fullName', ''))
        c.drawCentredString(right_x, sig_y - 70, secretary.get('fullName', ''))
        
        # Garis Bawah Nama
        c.setLineWidth(0.5)
        c.line(left_x - 40, sig_y - 72, left_x + 40, sig_y - 72)
        c.line(right_x - 40, sig_y - 72, right_x + 40, sig_y - 72)
        
        c.setFont("Times-Roman", 12)
        c.drawCentredString(left_x, sig_y - 85, chairman.get('NIM', ''))
        c.drawCentredString(right_x, sig_y - 85, secretary.get('NIM', ''))

        # Catatan Kaki
        c.setFont("Times-Italic", 9)
        c.drawString(margin + 20, 50, "*Surat ini dibuat dengan sistem informasi HIMATIKA dan ditandatangani secara elektronik.")
        c.drawString(margin + 20, 40, "*Verifikasi: https://himatika.itsnu.ac.id/signatures/scan")

        c.showPage() # End Page 1

        # --- PAGE 2 (LAMPIRAN) ---
        # Header Page 2 (Copy Header Page 1) - Disederhanakan untuk ringkas
        c.setFont("Times-Italic", 12)
        c.drawString(margin, height - 50, "Lampiran")
        c.setFont("Times-Roman", 12)
        c.drawString(margin + 20, height - 80, "Daftar keaktifan mahasiswa :")

        # Table Manual
        table_y = height - 110
        col_1_x = margin + 20
        col_2_x = margin + 300
        
        # Table Header
        c.setFont("Times-Bold", 12)
        c.drawString(col_1_x, table_y, "Kategori")
        c.drawString(col_2_x, table_y, "Jumlah")
        c.line(col_1_x, table_y - 5, width - margin - 20, table_y - 5)
        
        # Rows
        c.setFont("Times-Roman", 12)
        activities = point_data.get('activities', {})
        row_data = [
            ("Panitia Agenda", activities.get('agendas', {}).get('committees', 0)),
            ("Peserta Agenda", activities.get('agendas', {}).get('participants', 0)),
            ("Prestasi", activities.get('manualPoints', 0)),
            ("Proyek", activities.get('projects', 0)),
            ("Aspirasi", activities.get('aspirations', 0))
        ]
        
        curr_y = table_y - 25
        for label, val in row_data:
            c.drawString(col_1_x, curr_y, label)
            c.drawString(col_2_x, curr_y, str(val))
            curr_y -= 20

        c.save()
        buffer.seek(0)
        
        # Upload ke R2
        filename = f"Surat Keterangan Aktif {member.get('NIM')} Semester {point_data.get('semester')}.pdf"
        # Gunakan path /documents/activiness-letter/
        r2_key = f"documents/activiness-letter/{filename}"
        
        public_url = upload_bytes_to_r2(buffer.getvalue(), "application/pdf", r2_key)
        
        return jsonify({
            "success": True,
            "url": public_url,
            "filename": filename,
            "no": doc_number
        })

    except Exception as e:
        print(f"Error generating activiness letter: {e}")
        return jsonify({"error": str(e)}), 500


# --- 5. GENERATE TICKET ---
# Replikasi: makeTicket di useMakeDocs.ts
@app.route('/api/pdf/ticket', methods=['POST'])
def generate_ticket():
    try:
        body = request.json
        agenda = body.get('agenda')
        participant = body.get('participant') # Object IParticipant atau ICommittee
        role = body.get('role', 'participant') # 'participant' | 'committee'

        if not agenda or not participant:
            return jsonify({"error": "Data incomplete"}), 400

        # Setup Custom Page Size (600x250 points)
        custom_size = (600, 250)
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=custom_size)
        width, height = custom_size
        margin = 20

        # --- Background Image (Faded) ---
        # Ambil gambar pertama agenda
        photos = agenda.get('photos', [])
        if photos and len(photos) > 0:
            bg_url = photos[0].get('image')
            try:
                # Untuk background opacity di reportlab agak tricky
                # Kita gambar normal lalu timpa dengan kotak putih transparan
                # Atau gunakan fillAlpha jika versi reportlab mendukung
                bg_img = ImageReader(bg_url)
                c.saveState()
                c.setFillAlpha(0.1) # Transparansi
                c.drawImage(bg_img, 0, 0, width=width, height=height)
                c.restoreState()
            except:
                pass

        # --- Content ---
        # 1. Header: Event Title
        c.setFont("Helvetica-Bold", 18)
        title = agenda.get('title', 'AGENDA').upper()
        # Simple truncate if too long
        if len(title) > 35: title = title[:32] + "..."
        c.drawString(margin, height - 40, title)

        # 2. Info Grid
        c.setFont("Times-Roman", 8)
        c.setFillColor(Color(0.5, 0.5, 0.5))
        c.drawString(margin, height - 70, "TANGGAL")
        c.drawString(margin, height - 105, "WAKTU")
        c.drawString(margin, height - 140, "LOKASI")

        c.setFont("Times-Bold", 10)
        c.setFillColor(black)
        
        # Format Date (Simple)
        start_date = agenda.get('date', {}).get('start')
        try:
            date_obj = datetime.datetime.fromisoformat(start_date.replace('Z', '+00:00'))
            date_str = date_obj.strftime("%A, %d %B %Y")
            time_str = date_obj.strftime("%H:%M WIB")
        except:
            date_str = start_date
            time_str = "-"
            
        c.drawString(margin, height - 82, date_str)
        c.drawString(margin, height - 117, time_str)
        c.drawString(margin, height - 152, agenda.get('at', '-'))

        # 3. User Info (Bottom Left)
        # Ambil nama member
        member = participant.get('member')
        full_name = "Peserta"
        if isinstance(member, dict):
            full_name = member.get('fullName', 'Peserta')
        elif participant.get('guest'):
            full_name = participant.get('guest', {}).get('fullName', 'Tamu')
        
        # Role Badge
        role_label = "PANITIA" if role == 'committee' else "PESERTA"
        role_color = Color(0.8, 0, 0) if role == 'committee' else Color(0, 0.5, 0)
        
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(role_color)
        c.drawString(margin, 85, role_label)

        c.setFont("Times-Bold", 12)
        c.setFillColor(black)
        c.drawString(margin, 70, full_name)

        # Job Title (if committee)
        if role == 'committee':
            job = participant.get('job', '-')
            c.setFont("Times-Roman", 10)
            c.setFillColor(Color(0.3, 0.3, 0.3))
            c.drawString(margin, 58, job)

        # ID
        p_id = participant.get('_id', '-')
        c.setFont("Times-Roman", 8)
        c.setFillColor(Color(0.4, 0.4, 0.4))
        c.drawString(margin, 45, f"ID: {str(p_id)[-8:].upper()}")

        # 4. QR Code (Right Side)
        qr_size = 140
        qr_x = width - margin - qr_size
        qr_y = (height - qr_size) / 2
        
        # Generate QR Data
        qr_payload = {
            "a": agenda.get('_id'),
            "t": 'c' if role == 'committee' else 'p'
        }
        if role == 'committee':
            qr_payload['c'] = participant.get('_id')
        else:
            qr_payload['p'] = participant.get('_id')
            
        qr = qrcode.QRCode(box_size=10, border=0)
        qr.add_data(str(qr_payload)) # JSON stringify ala python
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="transparent")
        
        # Gambar QR ke Canvas
        # Convert PIL image to ReportLab Image
        img_byte_arr = io.BytesIO()
        qr_img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        
        # Border Dashed
        c.setDash(5, 5)
        c.setLineWidth(1)
        c.setStrokeColor(Color(0.7, 0.7, 0.7))
        c.rect(qr_x - 5, qr_y - 5, qr_size + 10, qr_size + 10)
        c.setDash([]) # Reset dash
        
        c.drawImage(ImageReader(img_byte_arr), qr_x, qr_y, width=qr_size, height=qr_size, mask='auto')
        
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(black)
        c.drawCentredString(qr_x + qr_size/2, qr_y - 20, "SCAN SAAT MASUK")

        c.save()
        buffer.seek(0)

        # Return File Binary (Downloadable)
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"Tiket-{role}-{full_name.split()[0]}.pdf"
        )

    except Exception as e:
        print(f"Error generating ticket: {e}")
        return jsonify({"error": str(e)}), 500

# --- 6. SIGNATURE OVERLAY (AUTO-DETECT LOCATION) ---
@app.route('/api/sign/process', methods=['POST'])
def process_sign_overlay():
    try:
        body = request.json
        pdf_url = body.get('pdf')
        output_path = body.get('outputBlobPath')
        qr_value = body.get('qrValue')
        
        # Opsional: User bisa kirim lokasi manual ATAU teks yang mau dicari
        manual_locations = body.get('locations', []) 
        search_text = body.get('searchText') # Misal: "/12345signature/"

        if not all([pdf_url, output_path, qr_value]):
            return jsonify({"error": "Missing required parameters"}), 400

        # 1. Download PDF ke Memory
        print(f"Downloading PDF from: {pdf_url}")
        response = requests.get(pdf_url)
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch PDF"}), 400
            
        pdf_bytes = response.content
        input_pdf_stream = io.BytesIO(pdf_bytes)

        # 2. Deteksi Lokasi Teks (Jika searchText ada)
        target_locations = manual_locations
        
        if search_text:
            print(f"Searching for text: {search_text}")
            # Buka dengan PyMuPDF untuk searching (sangat cepat & akurat)
            doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            detected_locs = []
            for i, page in enumerate(doc_fitz):
                # Cari semua instance teks
                text_instances = page.search_for(search_text)
                
                for rect in text_instances:
                    # rect adalah [x0, y0, x1, y1] (Top-Left system)
                    detected_locs.append({
                        "page": i + 1,      # 1-based index agar konsisten
                        "x": rect.x0,
                        "y": rect.y0,       # Top-Left Y
                        "width": rect.width,
                        "height": rect.height
                    })
            
            doc_fitz.close()
            
            # Gabungkan dengan manual locations jika ada
            target_locations.extend(detected_locs)

        if not target_locations:
            return jsonify({"error": "No signature location found. Please check searchText or locations."}), 400

        # 3. Proses Overlay (Menggunakan pypdf & reportlab seperti sebelumnya)
        # Kita gunakan pypdf untuk write agar struktur PDF aman
        reader = PdfReader(input_pdf_stream)
        writer = PdfWriter()

        # Generate QR Code (Sekali saja)
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, border=1, box_size=10)
        qr.add_data(qr_value)
        qr.make(fit=True)
        qr_pil_img = qr.make_image(fill_color="black", back_color="transparent")
        
        qr_byte_arr = io.BytesIO()
        qr_pil_img.save(qr_byte_arr, format='PNG')
        qr_byte_arr.seek(0)
        qr_image_reader = ImageReader(qr_byte_arr)

        # Kelompokkan lokasi per halaman
        locs_by_page = {}
        for loc in target_locations:
            p_num = int(loc.get('page', 1)) - 1
            if p_num not in locs_by_page:
                locs_by_page[p_num] = []
            locs_by_page[p_num].append(loc)

        for i, page in enumerate(reader.pages):
            if i in locs_by_page:
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)

                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(page_width, page_height))

                for loc in locs_by_page[i]:
                    # Jika hasil detect, ukurannya mungkin kecil (seukuran teks)
                    # Kita mungkin ingin QR-nya agak besar, misal fix 100x100 atau menyesuaikan
                    
                    x = float(loc.get('x', 0))
                    y_top = float(loc.get('y', 0))
                    
                    # Logic: Gunakan width/height dari deteksi, atau default 100 jika manual
                    # Jika dari deteksi, biasanya kita ingin menimpa teks signature dengan QR
                    w = float(loc.get('width', 100))
                    h = float(loc.get('height', 100))
                    
                    # Opsional: Jika deteksi teks biasanya kecil (tinggi font 12pt), 
                    # kita mungkin ingin memperbesar QR agar bisa discan (min 50-60pt)
                    if search_text and h < 50:
                        scale_ratio = 60 / h
                        w = w * scale_ratio
                        h = 60 # Set minimal height

                    # Konversi Koordinat (Top-Left ke Bottom-Left ReportLab)
                    y_bottom = page_height - (y_top + h) + 8 
                    
                    can.drawImage(qr_image_reader, x, y_bottom, width=w, height=h, mask='auto')

                can.save()
                packet.seek(0)
                
                overlay_pdf = PdfReader(packet)
                page.merge_page(overlay_pdf.pages[0])

            writer.add_page(page)

        # 4. Upload ke R2
        output_buffer = io.BytesIO()
        writer.write(output_buffer)
        output_buffer.seek(0)
        
        public_url = upload_bytes_to_r2(output_buffer.getvalue(), "application/pdf", output_path)

        return jsonify({
            "statusCode": 200,
            "statusMessage": "Sign processed successfully",
            "data": public_url
        })

    except Exception as e:
        print(f"Error processing sign: {e}")
        return jsonify({"statusCode": 500, "error": str(e)}), 500
if __name__ == '__main__':
    app.run(port=5000, debug=True)