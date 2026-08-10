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

from utils.helpers import month_to_roman, format_date_indo, draw_wrapped_text, draw_signature_box

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

        packets = []

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
                packets.append(packet)

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

