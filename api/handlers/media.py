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

bp = Blueprint('media', __name__)

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


def _compress_video_task(file_key: str, video_id: str, callback_url: str) -> None:
    """
    Background task: download raw video from R2, compress with FFmpeg, re-upload, notify backend.
    Uses context manager for temp files to prevent memory leaks.
    """
    s3 = get_s3_client()
    bucket_name = os.environ.get('R2_BUCKET_NAME', '')
    public_domain = os.environ.get('R2_PUBLIC_DOMAIN', '').rstrip('/')

    input_path = None
    output_path = None

    try:
        # 1. Extract extension from file_key if possible
        ext = os.path.splitext(file_key)[1]
        if not ext:
            ext = ".mp4"  # Default fallback

        # Download raw video from R2
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_input:
            input_path = tmp_input.name
            s3.download_fileobj(bucket_name, file_key, tmp_input)

        # 2. Compress with FFmpeg: Force output to WebM format
        output_path = input_path + "_compressed.webm"
        ffmpeg_cmd = [
            "ffmpeg", "-y",
            "-i", input_path,
            "-vf", "scale=-2:720",
            "-c:v", "libvpx-vp9",
            "-crf", "30",
            "-b:v", "0",
            "-c:a", "libopus",
            "-b:a", "128k",
            "-deadline", "realtime",
            "-cpu-used", "4",
            output_path
        ]
        content_type = "video/webm"
        base_key, _ = os.path.splitext(file_key)
        compressed_key = base_key.replace("raw_", "compressed_", 1) + ".webm"

        result = subprocess.run(ffmpeg_cmd, capture_output=True, text=True, timeout=600)
        if result.returncode != 0:
            print(f"FFmpeg stderr: {result.stderr}")
            raise RuntimeError(f"FFmpeg failed with code {result.returncode}")

        # 3. Upload compressed video to R2
        with open(output_path, "rb") as f:
            compressed_bytes = f.read()

        processed_url = upload_bytes_to_r2(compressed_bytes, content_type, compressed_key)

        # 4. Delete raw video from R2
        try:
            s3.delete_object(Bucket=bucket_name, Key=file_key)
        except Exception as del_err:
            print(f"Warning: Failed to delete raw video {file_key}: {del_err}")

        # 5. Notify backend via webhook
        requests.post(
            callback_url,
            json={
                "videoId": video_id,
                "status": "completed",
                "processedUrl": processed_url,
            },
            timeout=30,
        )
        print(f"Video {video_id} compressed successfully: {processed_url}")

    except Exception as e:
        print(f"Video compression failed for {video_id}: {e}")
        import traceback
        traceback.print_exc()

        try:
            requests.post(
                callback_url,
                json={
                    "videoId": video_id,
                    "status": "failed",
                },
                timeout=30,
            )
        except Exception as notify_err:
            print(f"Failed to notify backend of failure: {notify_err}")

    finally:
        for path in [input_path, output_path]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass



@bp.route('/api/media/compress-video', methods=['POST'])
def compress_video():
    try:
        body = request.json
        file_key: str = body.get('fileKey', '')
        video_id: str = body.get('videoId', '')
        callback_url: str = body.get('callbackUrl', '')

        if not all([file_key, video_id, callback_url]):
            return jsonify({"error": "Missing required parameters"}), 400

        thread = threading.Thread(
            target=_compress_video_task,
            args=(file_key, video_id, callback_url),
            daemon=True,
        )
        thread.start()

        return jsonify({
            "success": True,
            "message": "Video compression started",
            "videoId": video_id,
        }), 202

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

