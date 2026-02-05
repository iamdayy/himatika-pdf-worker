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