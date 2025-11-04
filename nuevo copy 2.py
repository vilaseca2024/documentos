from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import os
import textwrap

# ============================================================
# CONFIGURABLES (Datos del memorándum)
# ============================================================
fecha = "20/06/2025"
codigo_ref = "VLSCSA/RM/016/2025"
nombre_empleado = "Liliana Zambrana Paco"
nombre_remitente = "Edwing Mijael Delgadillo Navia"
cargo_remitente = "Jefe de Gestión y RRHH"
referencia = "I Entrega de credencial VILASECA"
lugar = "La Paz – Bolivia"
titulo_memorandum = "MEMORANDUM"

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
red_line_color = (165, 42, 42) # Color marrón/rojo oscuro
text_color = (0, 0, 0)

# --- Funciones para Cargar Fuentes ---
# Se mantiene la lógica para usar Times New Roman o un fallback
def get_truetype_or_default(base_paths, size):
    times_paths = ["C:/Windows/Fonts/times.ttf", "C:/Windows/Fonts/Times New Roman.ttf"]
    times_bold_paths = ["C:/Windows/Fonts/timesbd.ttf", "C:/Windows/Fonts/Times New Roman Bold.ttf"]

    if size == font_size_body or size == 26 or size == 30: # Fuentes normales
        paths = times_paths + base_paths
    else: # Fuentes negritas
        paths = times_bold_paths + base_paths

    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()

# Rutas de respaldo (DejaVu y Liberation)
base_font_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]
base_bold_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]

# Inicialización de fuentes
font_title = get_truetype_or_default(base_bold_paths, font_size_title)
font_title_one = get_truetype_or_default(base_bold_paths, font_size_title_one)
font_body = get_truetype_or_default(base_font_paths, font_size_body)
font_body_bold = get_truetype_or_default(base_bold_paths, font_size_body)
font_small = get_truetype_or_default(base_font_paths, 26)
font_website = get_truetype_or_default(base_font_paths, 30)


# ============================================================
# FUNCIÓN DE JUSTIFICACIÓN DE TEXTO (NUEVA)
# ============================================================
def justify_text(draw, line, x, y, width, font, color, line_height):
    """Dibuja una línea de texto justificada agregando espacio entre palabras."""
    words = line.split()
    if len(words) <= 1:
        # No justificar si es la última línea o solo una palabra
        draw.text((x, y), line, fill=color, font=font, align="left")
        return line_height + leading

    # 1. Calcular el ancho actual de la línea
    # Usamos textbbox para obtener el ancho
    bbox = draw.textbbox((0, 0), line.replace(" ", ""), font=font)
    text_width_no_spaces = bbox[2] - bbox[0]

    # 2. Calcular el espacio restante
    space_left = width - text_width_no_spaces

    # 3. Calcular el ancho de espacio necesario
    num_gaps = len(words) - 1
    if num_gaps > 0:
        space_width = space_left / num_gaps
    else:
        space_width = 0 # No hay espacios

    # 4. Dibujar la línea palabra por palabra
    current_x = x
    for i, word in enumerate(words):
        draw.text((current_x, y), word, fill=color, font=font)
        # Calcular el ancho de la palabra actual
        word_bbox = draw.textbbox((0, 0), word, font=font)
        word_width = word_bbox[2] - word_bbox[0]
        current_x += word_width # Avanzar la posición por el ancho de la palabra

        if i < num_gaps:
            current_x += space_width # Añadir el ancho de espacio calculado

    return line_height + leading

# ============================================================
# CREACIÓN DE IMAGEN BASE
# ============================================================
img = Image.new('RGB', (width, height), 'white')
draw = ImageDraw.Draw(img, 'RGBA')

# ============================================================
# LÍNEAS DE DISEÑO DE MARGEN (se mantienen)
# ============================================================
line1_x = 430
line2_x = 447
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


# ============================================================
# LOGOS Y ELEMENTOS (Sin cambios)
# ============================================================
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

# ============================================================
# MARCA DE AGUA (Sin cambios)
# ============================================================
try:
    marca_agua = Image.open('marca_agua.png').convert('RGBA')
    watermark_size = 1800
    wm_height = int(marca_agua.height * (watermark_size / marca_agua.width))
    marca_agua = marca_agua.resize((watermark_size, wm_height), Image.Resampling.LANCZOS)
    new_data = [(r, g, b, int(a * 0.1)) for r, g, b, a in marca_agua.getdata()]
    marca_agua.putdata(new_data)
    wm_x = (width - watermark_size) // 2 + 150
    wm_y = (height - wm_height) // 2 + 300
    img.paste(marca_agua, (wm_x, wm_y), marca_agua)
except Exception:
    pass

# ============================================================
# CONTACTOS (Sin cambios)
# ============================================================
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
draw.text((50, height - 130), "www.agencia-vilaseca.com", fill=text_color, font=font_website)

# ============================================================
# ENCABEZADO DEL MEMORÁNDUM (Ajustes de alineación y negritas)
# ============================================================
title_y = 580
vert_line_x = 1380          # Coordenada X de la línea vertical
header_bottom_y = 1350      # Coordenada Y de la doble línea horizontal
column_width = 600          # Ancho de las columnas (ajustado para que el centrado funcione)
left_col_x_start = 700      # Inicio de la columna izquierda
right_col_x = vert_line_x + 50 # Inicio de la columna derecha
left_col_x_center = left_col_x_start + (column_width / 2) # Centro de la columna izquierda
text_spacing = 70
text_start_y = 750

# 1. Título y Subrayado (Centrado)
center_offset = 200
draw.text((center_offset + width // 2, title_y), titulo_memorandum, fill=text_color, font=font_title_one, anchor="mm")
bbox = draw.textbbox((0, 0), titulo_memorandum, font=font_title_one)
title_w = bbox[2] - bbox[0]
title_h = bbox[3] - bbox[1]
underline_y = title_y + (title_h // 2) + 10
draw.line([(center_offset + width // 2 - title_w // 2, underline_y), (center_offset + width // 2 + title_w // 2, underline_y)], fill=text_color, width=5)

# 2. Bloque Izquierdo (Centrado en su columna)
# Usamos anchor="mm" para centrar cada línea en left_col_x_center
draw.text((left_col_x_center, text_start_y), "RECURSOS HUMANOS", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y + 1 * text_spacing), "VILASECA S.A.", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y + 2 * text_spacing), f"Fecha: {fecha}", fill=text_color, font=font_body_bold, anchor="mm")
draw.text((left_col_x_center, text_start_y + 3 * text_spacing), lugar, fill=text_color, font=font_body_bold, anchor="mm")

# 3. Bloque Derecho (Ajuste de Negritas)
draw.text((right_col_x, text_start_y), codigo_ref, fill=text_color, font=font_body_bold)

# Datos del Empleado y Remitente (Negritas en títulos)
current_right_y = text_start_y + 1 * text_spacing
draw.text((right_col_x, current_right_y), "Señor@:", fill=text_color, font=font_body_bold) # Debe ser negrita
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{nombre_empleado}", fill=text_color, font=font_body)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), "De:", fill=text_color, font=font_body_bold) # Debe ser negrita
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{nombre_remitente}", fill=text_color, font=font_body)
current_right_y += 1 * text_spacing
draw.text((right_col_x, current_right_y), f"{cargo_remitente}", fill=text_color, font=font_body_bold) # Debe ser negrita
current_right_y += 1 * text_spacing

# Referencia (Ref)
draw.text((right_col_x, header_bottom_y - 2 * text_spacing), f"Ref.: {referencia}", fill=text_color, font=font_body_bold) # Debe ser negrita

# 4. Línea Separadora Vertical
vert_line_start_y = title_y + title_h + 80
vert_line_end_y = header_bottom_y - 4
draw.line([(vert_line_x, vert_line_start_y), (vert_line_x, vert_line_end_y)], fill=text_color, width=2)
draw.line([(vert_line_x + 7, vert_line_start_y), (vert_line_x + 7, vert_line_end_y)], fill=text_color, width=2)
# 5. Doble Línea Separadora Horizontal (Full-Width)
draw.line([(650, header_bottom_y), (width - 150, header_bottom_y)], fill=text_color, width=2)
draw.line([(650, header_bottom_y + 7), (width - 150, header_bottom_y + 7)], fill=text_color, width=2)


# ============================================================
# TEXTO PRINCIPAL (JUSTIFICADO Y CON NEGRILLAS CORREGIDAS)
# ============================================================
body_text_x_start = left_col_x_start # Posición X para inicio de texto
body_text_y_start = header_bottom_y + 60 # Posición Y alineada a la línea divisoria

# --- Contenido completo y detallado del memorándum ---
paragraphs_to_wrap = [
    f"Estimad@ {nombre_empleado}:", # Primera línea debe tener Estimad@ en negrita
    "Mediante la presente, se le hace entrega de su PRIMERA credencial, que acredita que usted realiza funciones en nuestra agencia. El mismo tiene una vigencia de 3 años desde la fecha de emisión de la misma o hasta la culminación de la relación laboral o de pasantía.",
    "Hacerle conocer que a partir de la fecha debe portar todos los días esta credencial, que está bajo su entera responsabilidad, y en caso de una pérdida o daño en uno de ellos deberá procederse de acuerdo al procedimiento de Recursos Humanos.",
    "En caso que usted deje de pertenecer a la empresa, la credencial deberá ser devuelta al área de Recursos Humanos (junto a la cinta y porta credencial acrílica) o cubrir con el costo de la misma.",
    "Es bueno recalcarle, que las visitas de clientes y proveedores, requerirán una credencial de visita para ingreso a nuestras oficinas." # Este párrafo va antes de la firma
]

# Items de costo y firma
costos_centrados = [
    "Credencial 25 Bs",
    "Cinta porta credencial 15 Bs",
    "Protector porta credencial 10 Bs"
]

firma = [
    "Atentamente:",
    "",
    "",
    f"{nombre_remitente}",
    f"{cargo_remitente}"
]

pie_pagina_izquierda = [
    "EMD/",
    "C.c. File personal"
]

# Configuración de salto de línea
letter_block_width = width - body_text_x_start - 150 # Ancho total de la columna de texto (hasta el margen derecho)
line_height = font_body.getbbox("A")[3] - font_body.getbbox("A")[1]
current_y = body_text_y_start

# --- 1. Dibujar Párrafo 1 (con negrita y Justificado) ---
# Separar "Estimad@" y "Liliana Zambrana Paco"
draw.text((body_text_x_start, current_y), "Estimad@", fill=text_color, font=font_body_bold, align="left")
bbox_estimado = draw.textbbox((0, 0), "Estimad@", font=font_body_bold)
width_estimado = bbox_estimado[2] - bbox_estimado[0]

# Continuar con el nombre sin negrita
draw.text((body_text_x_start + width_estimado + 10, current_y), f"{nombre_empleado}:", fill=text_color, font=font_body, align="left")
current_y += line_height + leading # Siguiente línea

# El resto de párrafos Justificados
for i, paragraph in enumerate(paragraphs_to_wrap[1:]):
    wrapped_lines = textwrap.wrap(paragraph, width=100) # Usamos un número alto para que textwrap divida
    
    # Manejar el último párrafo (índice 3 en la lista ajustada [1:]) separadamente
    if i == 3:
        current_y += line_height * 2 # Espacio adicional antes de "Es bueno recalcarle"
        
        # Este párrafo NO se justifica, se alinea a la izquierda.
        for line in wrapped_lines:
            draw.text((body_text_x_start, current_y), line, fill=text_color, font=font_body, align="left")
            current_y += line_height + leading
        break # Salimos del bucle después de manejar el último párrafo antes de costos

    # Justificación para todos los párrafos intermedios
    for j, line in enumerate(wrapped_lines):
        # Si NO es la última línea del párrafo, justificar
        if j < len(wrapped_lines) - 1:
            justify_text(draw, line, body_text_x_start, current_y, letter_block_width, font_body, text_color, line_height)
        else:
            # Última línea del párrafo: alinear izquierda
            draw.text((body_text_x_start, current_y), line, fill=text_color, font=font_body, align="left")
        
        current_y += line_height + leading
    current_y += line_height # Espacio entre párrafos

# --- 2. Dibujar Costos (Centrados en la columna de texto) ---
center_x = body_text_x_start + (letter_block_width / 2)
current_y += line_height # Espacio adicional antes de los costos

for costo in costos_centrados:
    draw.text((center_x, current_y), costo, fill=text_color, font=font_body, anchor="mm")
    current_y += line_height + leading

current_y += line_height * 2 # Espacio adicional después de los costos


# --- 3. Dibujar Firma (Centrada en la columna) ---
# Reajustamos la firma con las negritas correctas según la imagen (Nombre y Cargo en Negrita, Atentamente en Negrita)
current_y += line_height # Espacio antes de Atentamente

# Atentamente: (Negrita y a la izquierda de la columna)
draw.text((body_text_x_start, current_y), "Atentamente:", fill=text_color, font=font_body_bold, align="left")
current_y += line_height + leading * 4 # Espacio para la firma (líneas vacías)

# Nombre Centrado y en Negrita
draw.text((center_x, current_y), f"{nombre_remitente}", fill=text_color, font=font_body_bold, anchor="mm")
current_y += line_height + leading

# Cargo Centrado y en Negrita
draw.text((center_x, current_y), f"{cargo_remitente}", fill=text_color, font=font_body_bold, anchor="mm")
current_y += line_height + leading

# --- 4. Dibujar Pie de Página Izquierdo ---
current_y += line_height * 2 # Espacio después de la firma

for line in pie_pagina_izquierda:
    draw.text((body_text_x_start, current_y), line, fill=text_color, font=font_body, align="left")
    current_y += line_height + leading

# ============================================================
# EXPORTAR A PDF (Sin cambios)
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
