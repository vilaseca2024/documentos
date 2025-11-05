from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import os

# ============================================================
# CONFIGURABLES (Datos del memorándum)
# ============================================================
fecha = "20/06/2025"
codigo_ref = "VLSCSA/RM/016/2025"
nombre_empleado = "Liliana Zambrana Paco"
nombre_remitente = "Edwing Mijael Delgadillo Navia"
cargo_remitente = "Jefe de Gestión y RRHH"
 
referencia = "Entrega Credencial Personal \nAcreditado ante la AN"
lugar = "La Paz – Bolivia"
titulo_memorandum = "MEMORANDUM"
fecha_inicio = "20/06/2025"
fecha_fin = "20/06/2025"

# ============================================================
# CONFIGURACIÓN DE DOCUMENTO
# ============================================================
os.makedirs('output', exist_ok=True)

width = 2550  # 8.5 inches * 300 DPI
height = 3300  # 11 inches * 300 DPI

font_size_title = 35
font_size_title_one = 60
font_size_body = 50
leading = 6
leading2 = 14
red_line_color = (165, 42, 42) # Color marrón/rojo oscuro
text_color = (0, 0, 0)

# --- Funciones para Cargar Fuentes (Mejorado para forzar Times New Roman) ---
# Intentamos cargar Times New Roman. Si no existe, usamos fallbacks genéricos.
def get_truetype_or_default(base_paths, size, is_bold=False):
    if is_bold:
        times_paths = ["C:/Windows/Fonts/timesbd.ttf", "/usr/share/fonts/msttcore/Times_New_Roman_Bold.ttf"]
    else:
        times_paths = ["C:/Windows/Fonts/times.ttf", "/usr/share/fonts/msttcore/Times_New_Roman.ttf"]
    for p in times_paths + base_paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default() # Fallback si nada funciona

base_font_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]
base_bold_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]

font_title = get_truetype_or_default(base_bold_paths, font_size_title, is_bold=True)
font_title_one = get_truetype_or_default(base_bold_paths, font_size_title_one, is_bold=True)
font_body = get_truetype_or_default(base_font_paths, font_size_body, is_bold=False)
font_body_bold = get_truetype_or_default(base_bold_paths, font_size_body, is_bold=True)
font_body_pie = get_truetype_or_default(base_font_paths, font_size_body, is_bold=False)
font_small = get_truetype_or_default(base_font_paths, 26, is_bold=False)
font_website = get_truetype_or_default(base_font_paths, 30, is_bold=False)


# ============================================================
# Helpers: wrap por píxeles y justificación mejorada
# ============================================================

bold_words = set([
    "PÉRDIDA", "O", "DETERIORO",
    "RD", "01-085-24", "01-003-24",
    "Clasificación", "Contravenciones", "Aduaneras",
    "LABORAL",
    "Credencial,", "Porta", "credencial", "y", "cinta", "OEA.",
    "AN"
])

def draw_paragraph_with_bold_justified(draw, text, start_x, start_y, max_width, font_normal, font_bold, color, line_height, leading):
    words = text.split()
    lines = []
    curr = ""
    for w in words:
        test = curr + (" " if curr else "") + w
        w_px = draw.textbbox((0,0), test, font=font_normal)[2]
        if w_px <= max_width:
            curr = test
        else:
            lines.append(curr)
            curr = w
    lines.append(curr)

    y = start_y
    for i, line in enumerate(lines):
        parts = line.split()
        total_words_width = 0
        parts_widths = []
        for part in parts:
            f = font_bold if part.strip(",.").upper() in bold_words else font_normal
            pw = draw.textbbox((0,0), part, font=f)[2]
            parts_widths.append((part, f, pw))
            total_words_width += pw

        gaps = len(parts) - 1
        space_w = draw.textbbox((0,0), " ", font=font_normal)[2]

        if gaps > 0 and i != len(lines) - 1:
            extra = (max_width - (total_words_width + space_w*gaps)) / gaps
        else:
            extra = 0

        x = start_x
        for idx, (part, f, pw) in enumerate(parts_widths):
            draw.text((x, y), part, fill=color, font=f)
            x += pw
            if idx < gaps:
                x += space_w + extra

        y += line_height + leading

    return y



def wrap_text_by_pixels(draw, text, font, max_width):
    """Rompe el texto en líneas que no excedan max_width (en píxeles)."""
    words = text.split()
    if not words:
        return []
    lines = []
    current = words[0]
    for w in words[1:]:
        test = current + " " + w
        bbox = draw.textbbox((0,0), test, font=font)
        w_px = bbox[2] - bbox[0]
        if w_px <= max_width:
            current = test
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return lines

def justify_text(draw, line, x, y, width_px, font, color):
    """
    Dibuja una línea justificada en el ancho width_px.
    Calcula el ancho de cada palabra y el ancho de un espacio simple.
    """
    words = line.split()
    if len(words) == 0:
        return
    if len(words) == 1:
        draw.text((x, y), line, fill=color, font=font, align="left")
        return

    # medir ancho de cada palabra
    word_widths = []
    for w in words:
        bbox = draw.textbbox((0,0), w, font=font)
        word_widths.append(bbox[2] - bbox[0])
    # ancho del espacio simple
    sp_bbox = draw.textbbox((0,0), " ", font=font)
    space_w = sp_bbox[2] - sp_bbox[0]

    total_words = sum(word_widths)
    num_gaps = len(words) - 1
    total_natural_spaces = space_w * num_gaps
    extra_space_total = max(0, width_px - (total_words + total_natural_spaces))
    extra_per_gap = extra_space_total / num_gaps if num_gaps > 0 else 0

    current_x = x
    for i, w in enumerate(words):
        draw.text((current_x, y), w, fill=color, font=font)
        current_x += word_widths[i]
        if i < num_gaps:
            current_x += space_w + extra_per_gap

# ============================================================
# CREACIÓN DE IMAGEN BASE (sin cambios estructurales)
# ============================================================
img = Image.new('RGB', (width, height), 'white')
draw = ImageDraw.Draw(img, 'RGBA')

# Líneas de margen
line1_x = 445
line2_x = 462
line1_width = 10
line2_width = 3
horizontal_line_start_y = 530
horizontal_line_end_y = height - 100

draw.rectangle([line1_x, horizontal_line_start_y, line1_x + line1_width, horizontal_line_end_y], fill=red_line_color)
draw.rectangle([line2_x, horizontal_line_start_y, line2_x + line2_width, horizontal_line_end_y], fill=red_line_color)

horizontal_line_start_x = 1270
horizontal_line_end_x = width - 150
horizontal_line1_y = 305
horizontal_line2_y = 315
draw.rectangle([horizontal_line_start_x, horizontal_line1_y, horizontal_line_end_x, horizontal_line1_y + 3], fill=red_line_color)
draw.rectangle([horizontal_line_start_x, horizontal_line2_y, horizontal_line_end_x, horizontal_line2_y + 10], fill=red_line_color)

# Logos/elementos (intento cargar, si no, texto marcador)
logo_left_margin = 100
logo_left_margin2 = 70
try:
    logo = Image.open('logo.png')
    logo_width = 1200
    logo_height = int(logo.height * (logo_width / logo.width))
    logo = logo.resize((logo_width, logo_height), Image.Resampling.LANCZOS)
    img.paste(logo, (logo_left_margin, 100), logo if logo.mode == 'RGBA' else None)
except Exception:
    draw.text((logo_left_margin + 225, 220), "LOGO.PNG", fill='red', font=font_title, anchor="mm")

for name, y_pos, color in [("hexagono", 700, "blue"), ("rectangulo", 1150, "green"), ("triangulo", 1600, "orange")]:
    try:
        image = Image.open(f'{name}.png')
        width_img = 320 if name == "hexagono" else (250 if name == "rectangulo" else 300)
        height_img = int(image.height * (width_img / image.width))
        image = image.resize((width_img, height_img), Image.Resampling.LANCZOS)
        img.paste(image, (logo_left_margin2 + 30 if name == "rectangulo" else logo_left_margin2 , y_pos), image if image.mode == 'RGBA' else None)
    except Exception:
        draw.text((logo_left_margin2 + 60, y_pos + 100), name.upper(), fill=color, font=font_small, anchor="mm")

# Marca de agua (opcional)
try:
    marca_agua = Image.open('marca_agua.png').convert('RGBA')
    watermark_size = 1800
    wm_height = int(marca_agua.height * (watermark_size / marca_agua.width))
    marca_agua = marca_agua.resize((watermark_size, wm_height), Image.Resampling.LANCZOS)
    new_data = [(r, g, b, int(a )) for r, g, b, a in marca_agua.getdata()] 
    marca_agua.putdata(new_data)
    wm_x = (width - watermark_size) // 2 + 250
    wm_y = (height - wm_height) // 2 + 300
    img.paste(marca_agua, (wm_x, wm_y), marca_agua)
except Exception:
    pass

# Contactos (sin cambios)
la_paz_y = 2400
margin_left_letras = 45
draw.text((margin_left_letras, la_paz_y), "LA PAZ", fill=text_color, font=font_title)
draw.text((margin_left_letras, la_paz_y + 50), "Calle Mercado N°1328", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 80), "Entre Loayza y Colon", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 110), "Edif. Mariscal Ballivian", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 140), "Piso 7 Of. 702-706", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 170), "Central Piloto (+591)-(2)2202077", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 200), "Celular Piloto (+591)-67005008", fill=text_color, font=font_small)
draw.text((margin_left_letras, la_paz_y + 230), "Casilla de Correo 1482", fill=text_color, font=font_small)

sc_y = 2800
draw.text((margin_left_letras, sc_y), "SANTA CRUZ", fill=text_color, font=font_title)
draw.text((margin_left_letras, sc_y + 50), "Calle Marayaú #2465", fill=text_color, font=font_small)
draw.text((margin_left_letras, sc_y + 80), "Entre Beni y Alemana", fill=text_color, font=font_small)
draw.text((margin_left_letras, sc_y + 110), "Paralela a la Av. Los Cusis", fill=text_color, font=font_small)
draw.text((margin_left_letras, sc_y + 140), "Telf.: (+591)-67005008", fill=text_color, font=font_small)
draw.text((50, height - 140), "www.agencia-vilaseca.com", fill=text_color, font=font_website)

# ============================================================
# ENCABEZADO DEL MEMORÁNDUM (igual que antes)
# ============================================================
title_y = 680
vert_line_x = 1380
header_bottom_y = 1450
left_col_x = 600
column_width = 600
left_col_x_start = 700
left_col_x_center = left_col_x_start + (column_width / 2)
right_col_x = vert_line_x + 50
text_spacing = 70
text_start_y = 850
text_start_y2 = 1000

center_offset = 200
draw.text((center_offset + width // 2, title_y), titulo_memorandum, fill=text_color, font=font_title_one, anchor="mm")
bbox = draw.textbbox((0, 0), titulo_memorandum, font=font_title_one)
title_w = bbox[2] - bbox[0]
title_h = bbox[3] - bbox[1]
underline_y = title_y + (title_h // 2) + 10
draw.line([(center_offset + width // 2 - title_w // 2, underline_y), (center_offset + width // 2 + title_w // 2, underline_y)], fill=text_color, width=5)

# Bloque izquierdo
draw.text((left_col_x_center, text_start_y2), "RECURSOS HUMANOS", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y2 + 1 * text_spacing), "VILASECA S.A.", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y2 + 2 * text_spacing), f"Fecha: {fecha}", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y2 + 3 * text_spacing), lugar, fill=text_color, font=font_body_bold, anchor="mm")

# Bloque derecho
draw.text((right_col_x, text_start_y), codigo_ref, fill=text_color, font=font_body_bold)
current_right_y = text_start_y + 1 * text_spacing
draw.text((right_col_x, current_right_y), "Señor/a:", fill=text_color, font=font_body_bold)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{nombre_empleado}", fill=text_color, font=font_body)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), "De:", fill=text_color, font=font_body_bold)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{nombre_remitente}", fill=text_color, font=font_body)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{cargo_remitente}", fill=text_color, font=font_body_bold)
current_right_y += 1 * text_spacing

draw.text((right_col_x, header_bottom_y - 2 * text_spacing), f"Ref.: {referencia}", fill=text_color, font=font_body_bold)

vert_line_start_y = title_y + title_h + 80
vert_line_end_y = header_bottom_y - 4
draw.line([(vert_line_x, vert_line_start_y), (vert_line_x, vert_line_end_y)], fill=text_color, width=2)
draw.line([(vert_line_x + 7, vert_line_start_y), (vert_line_x + 7, vert_line_end_y)], fill=text_color, width=2)
draw.line([(600, header_bottom_y), (width - 150, header_bottom_y)], fill=text_color, width=2)
draw.line([(600, header_bottom_y + 7), (width - 150, header_bottom_y + 7)], fill=text_color, width=2)

# ============================================================
# TEXTO PRINCIPAL (JUSTIFICADO Y CON VIÑETAS - mejora)
# ============================================================
body_text_x_start = left_col_x
body_text_y_start = header_bottom_y + 60

fecha_inicio_credencial = "21 de octubre del 2025"
fecha_fin_credencial   = "21 de octubre del 2027"

closing_paragraphs = [
    f"Mediante la presente, se le hace la entrega de la credencial emitida por la AN de personal acreditado para realizar trámites inherentes a la gestión de despachos en las diferentes aduanas; la misma tendrá validez a partir del {fecha_inicio_credencial}, hasta el {fecha_fin_credencial}.\n\n"
    "Informarle que a partir de la fecha debe portar todos los días esta credencial, que está bajo su entera responsabilidad. En caso de PÉRDIDA O DETERIORO se procederá de acuerdo a la RD 01-085-24 la cual indica que en caso de reposición será pasible a la aplicación de sanciones establecidas en la RD 01-003-24 Clasificación de Contravenciones Aduaneras, independientemente de las acciones legales que correspondan por la mala utilización de la credencial, mismo que será asumido por su persona.\n\n"
    "En caso que usted deje de mantener una relación LABORAL con VILASECA S.A., la credencial deberá ser devuelta al área de Recursos Humanos para el envió a la Aduana Nacional para su destrucción, en la devolución debe entregar: Credencial, Porta credencial y cinta OEA.\n\n"
    "De la misma forma en caso de estar cerca el vencimiento de la misma informar a Recursos Humanos para tramitar la renovación ante la AN con una semana de anticipación.\n\n"
]


 

letter_block_width = width - body_text_x_start - 150
line_height = font_body.getbbox("A")[3] - font_body.getbbox("A")[1]
current_y = body_text_y_start

# Saludo
draw.text((body_text_x_start, current_y), "Estimado/a ", fill=text_color, font=font_body, align="left")
bbox_estimado = draw.textbbox((0, 0), "Estimado/a ", font=font_body)
width_estimado = bbox_estimado[2] - bbox_estimado[0]
draw.text((body_text_x_start + width_estimado + 10, current_y), f"{nombre_empleado}:", fill=text_color, font=font_body, align="left")
current_y += line_height + leading * 10



# Párrafos de cierre
# Párrafos de cierre
for p in closing_paragraphs:
    # Divide el texto en sub-párrafos usando \n\n
    sub_paragraphs = p.split("\n\n")
    for sub_p in sub_paragraphs:
        current_y = draw_paragraph_with_bold_justified(
            draw, sub_p,
            body_text_x_start, current_y,
            letter_block_width,
            font_body, font_body_bold,
            text_color, line_height, leading2
        )
        # Añade un salto de línea extra entre sub-parrafos
        current_y += line_height * 1  # puedes ajustar la separación



current_y += line_height - 40

# Firma
draw.text((body_text_x_start, current_y), "Atentamente:", fill=text_color, font=font_body, align="left")
current_y += line_height + leading2 * 15

center_x = body_text_x_start + (letter_block_width / 2)
draw.text((center_x, current_y), f"{nombre_remitente}", fill=text_color, font=font_body, anchor="mm")
current_y += line_height + leading2
draw.text((center_x, current_y), f"{cargo_remitente}", fill=text_color, font=font_body_bold, anchor="mm")
current_y += line_height + leading2

# Pie de página izquierdo
current_y += line_height * 2
pie_pagina_izquierda = [
    "EMD/",
    "C.c. File personal"
]
for line in pie_pagina_izquierda:
    draw.text((body_text_x_start, current_y), line, fill=text_color, font=font_small, align="left")
    current_y += line_height + leading

# ============================================================
# EXPORTAR A PDF
# ============================================================
temp_png_path = 'output/letterhead_vilaseca_temp.png'
img.save(temp_png_path, 'PNG', dpi=(300, 300))

pdf_path = 'output/letterhead_vilaseca.pdf'
c = canvas.Canvas(pdf_path, pagesize=letter)
letter_width, letter_height = letter
img_reader = ImageReader(temp_png_path)
c.drawImage(img_reader, 0, 0, width=letter_width, height=letter_height, preserveAspectRatio=True)
c.save()
os.remove(temp_png_path)

print(f"✓ Memorándum PDF generado correctamente: {pdf_path}")
