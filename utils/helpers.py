def flatten_object(obj, prefix=""):
    result = {}
    key_not_allow = ["_id", "id", "__v", "createdAt", "updatedAt"]
    
    for key, value in obj.items():
        if key in key_not_allow:
            continue
            
        if isinstance(value, dict) and value is not None:
            # Rekursif flatten
            flat_child = flatten_object(value)
            result.update(flat_child)
        else:
            result[key] = value
            
    return result

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
    Mengembalikan koordinat Y terakhir setelah teks selesai ditulis.
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth
    
    if line_height is None:
        line_height = font_size + 2

    words = text.split(" ")
    current_line = []
    current_y = y
    
    canvas_obj.setFont(font_name, font_size)

    for word in words:
        # Cek panjang baris jika kata ini ditambahkan
        test_line = " ".join(current_line + [word])
        width = stringWidth(test_line, font_name, font_size)
        
        if width <= max_width:
            current_line.append(word)
        else:
            # Tulis baris saat ini
            canvas_obj.drawString(x, current_y, " ".join(current_line))
            current_line = [word] # Mulai baris baru dengan kata saat ini
            current_y -= line_height # Turun ke bawah

    # Tulis sisa kata di baris terakhir
    if current_line:
        canvas_obj.drawString(x, current_y, " ".join(current_line))
    
    return current_y # Kembalikan posisi Y terakhir

def draw_signature_box(c, x, y_top, w, h, p_h,
                        sig_name='', sig_as='', sig_img=None, overlap=False):
    from reportlab.lib.colors import Color
    role_h   = h * 0.15   # jabatan di atas
    name_h   = h * 0.20   # nama di bawah
    img_h    = h - role_h - name_h   # sisa → gambar

    y_bot    = p_h - (y_top + h)
    name_y0  = y_bot               # bawah area nama
    name_y1  = y_bot + name_h      # batas atas nama / batas bawah garis
    img_y0   = name_y1             # bawah area gambar (tanpa overlap)
    img_y1   = img_y0 + img_h      # atas area gambar / batas bawah jabatan
    role_y0  = img_y1

    c.saveState()

    if sig_as:
        role_fs = max(12, min(14, role_h * 0.60))
        c.setFont('Helvetica', role_fs)
        c.setFillColor(Color(0.2, 0.2, 0.2))
        role_text_y = role_y0 + (role_h - role_fs) / 2
        c.drawCentredString(x + w / 2, role_text_y, sig_as[:50])

    c.setStrokeColor(Color(0.5, 0.5, 0.5))
    c.setLineWidth(0.3)
    c.line(x, img_y1, x + w, img_y1)

    c.setStrokeColor(Color(0.5, 0.5, 0.5))
    c.setLineWidth(0.5)
    c.line(x, name_y1, x + w, name_y1)

    if sig_name:
        name_fs = max(12, min(14, name_h * 0.60))
        c.setFont('Helvetica-Bold', name_fs)
        c.setFillColor(Color(0.1, 0.1, 0.1))
        name_text_y = name_y0 + (name_h - name_fs) / 2
        c.drawCentredString(x + w / 2, name_text_y, sig_name[:40])

    if sig_img is not None:
        if overlap:
            overlap_px  = name_h * 0.50
            draw_y      = img_y0 - overlap_px
            draw_h      = img_h + overlap_px
        else:
            draw_y  = img_y0
            draw_h  = img_h
        side     = min(w * 0.92, draw_h)
        draw_x   = x + (w - side) / 2
        center_y = draw_y + (draw_h - side) / 2
        c.drawImage(sig_img, draw_x, center_y,
                    width=side, height=side,
                    mask='auto', preserveAspectRatio=False)

    c.restoreState()