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
red_line_color = (165, 42, 42)
text_color = (0, 0, 0)

def load_font_prefer(paths_sizes):
    for path, size in paths_sizes:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()

base_font_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "C:/Windows/Fonts/arial.ttf"
]
base_bold_paths = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "C:/Windows/Fonts/arialbd.ttf"
]

def get_truetype_or_default(base_paths, size):
    for p in base_paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()

font_title = get_truetype_or_default(base_bold_paths, font_size_title)
font_title_one = get_truetype_or_default(base_bold_paths, font_size_title_one)
font_body = get_truetype_or_default(base_font_paths, font_size_body)
font_small = get_truetype_or_default(base_font_paths, 26)
font_website = get_truetype_or_default(base_font_paths, 30)

# ============================================================
# CREACIÓN DE IMAGEN BASE
# ============================================================
img = Image.new('RGB', (width, height), 'white')
draw = ImageDraw.Draw(img, 'RGBA')

# ============================================================
# LÍNEAS DE DISEÑO
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
# LOGOS Y ELEMENTOS
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
        img.paste(image, (logo_left_margin2, y_pos), image if image.mode == 'RGBA' else None)
    except Exception:
        draw.text((logo_left_margin2 + 60, y_pos + 100), name.upper(), fill=color, font=font_small, anchor="mm")

# ============================================================
# MARCA DE AGUA
# ============================================================
try:
    marca_agua = Image.open('marca_agua.png').convert('RGBA')
    watermark_size = 1800
    wm_height = int(marca_agua.height * (watermark_size / marca_agua.width))
    marca_agua = marca_agua.resize((watermark_size, wm_height), Image.Resampling.LANCZOS)
    new_data = [(r, g, b, int(a )) for r, g, b, a in marca_agua.getdata()]
    marca_agua.putdata(new_data)
    wm_x = (width - watermark_size) // 2 + 150
    wm_y = (height - wm_height) // 2 + 300
    img.paste(marca_agua, (wm_x, wm_y), marca_agua)
except Exception:
    pass

# ============================================================
# CONTACTOS
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
# ENCABEZADO DEL MEMORÁNDUM
# ============================================================
draw.text((1330, 580), titulo_memorandum, fill=text_color, font=font_title_one, anchor="mm")
draw.text((700, 750), "RECURSOS HUMANOS", fill=text_color, font=font_body)
draw.text((700, 820), "VILASECA S.A.", fill=text_color, font=font_body)
draw.text((700, 890), f"Fecha: {fecha}", fill=text_color, font=font_body)
draw.text((700, 960), lugar, fill=text_color, font=font_body)

# Bloque derecho
draw.text((1500, 750), codigo_ref, fill=text_color, font=font_body)
draw.text((1500, 850), f"Señor@:\n{nombre_empleado}\n\nDe:\n{nombre_remitente}\n{cargo_remitente}", fill=text_color, font=font_body, spacing=5)
draw.text((1500, 1150), f"Ref.: {referencia}", fill=text_color, font=font_body)

# ============================================================
# TEXTO PRINCIPAL
# ============================================================
sample_letter = (
    f"Estimad@ {nombre_empleado}:\n\n"
    "Mediante la presente, se le hace entrega de su PRIMERA credencial, que acredita que usted realiza funciones en nuestra agencia. "
    "El mismo tiene una vigencia de 3 años desde la fecha de emisión de la misma o hasta la culminación de la relación laboral o de pasantía.\n\n"
    "Hacerle conocer que a partir de la fecha debe portar todos los días esta credencial, que está bajo su entera responsabilidad, "
    "y en caso de una pérdida o daño en uno de ellos deberá procederse de acuerdo al procedimiento de Recursos Humanos.\n\n"
    "En caso que usted deje de pertenecer a la empresa, la credencial deberá ser devuelta al área de Recursos Humanos "
    "(junto a la cinta y porta credencial acrílica) o cubrir con el costo de la misma.\n\n"
    "Atentamente,\n\n"
    f"{nombre_remitente}\n"
    f"{cargo_remitente}\n"
)

letter_block_width = 800
center_x = width // 2 + 200
center_y = 1700

sample_char_width = font_body.getbbox("M")[2] - font_body.getbbox("M")[0]

approx_chars_per_line = max(80, letter_block_width // max(1, sample_char_width))
wrapped = []
for paragraph in sample_letter.split("\n\n"):
    wrapped.extend(textwrap.wrap(paragraph, width=approx_chars_per_line))
    wrapped.append("")
wrapped_text = "\n".join(wrapped).strip()

try:
    bbox = draw.multiline_textbbox((0, 0), wrapped_text, font=font_body, spacing=leading)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
except Exception:
    text_w = letter_block_width
    line_h = font_body.getsize("Ay")[1]
    text_h = len(wrapped_text.splitlines()) * (line_h + leading)

text_x = center_x - text_w // 2
text_y = center_y - text_h // 2
draw.multiline_text((text_x, text_y), wrapped_text, fill=text_color, font=font_body, spacing=leading, align="left")

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
