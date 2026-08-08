import os

with open('api/index.py', 'r') as f:
    lines = f.readlines()

def extract_lines(start, end):
    # lines are 1-indexed in our previous script output
    # start and end are 1-indexed
    # return the joined string
    return "".join(lines[start-1:end]).replace("@app.route", "@bp.route")

IMPORTS = """import sys
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
"""

helpers_lines = "".join(lines[45:165])

BLUEPRINTS = {
    'pdf': [
        (190, 274), # scan_qr
        (337, 387), # search_text_pdf
        (570, 1074), # generate_activiness_letter
        (1075, 1084), # rupiah_format helper
        (1085, 1306), # generate_ticket
        (1441, 1496), # certificate_preview
        (1498, 1666), # generate_certificate
    ],
    'sheet': [
        (275, 335), # import_generic_sheet
        (389, 468), # export_generic_sheet
    ],
    'sign': [
        (470, 567), # process_sign_overlay
    ],
    'tools': [
        (1308, 1338), # generate_qr_tool
        (1339, 1398), # compress_image_tool
        (1403, 1438), # upload_image_to_r2
    ],
    'media': [
        (1669, 1765), # _compress_video_task
        (1766, 1794), # compress_video
    ]
}

os.makedirs('api/handlers', exist_ok=True)

for bp_name, ranges in BLUEPRINTS.items():
    content = IMPORTS + f"\nbp = Blueprint('{bp_name}', __name__)\n\n" + helpers_lines + "\n"
    for r in ranges:
        content += extract_lines(r[0], r[1]) + "\n"
    
    with open(f'api/handlers/{bp_name}.py', 'w') as f:
        f.write(content)

print("Generated handler files.")
