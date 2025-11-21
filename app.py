from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from flask import send_from_directory
import os
from werkzeug.utils import secure_filename
import re
import uuid

# Color mappings for highlights
_colors = {
    'yellow': 'FFFF00',
    'green': '00FF00',  
    'blue': '00FFFF',
    'red': 'FF0000'
}

# Optional conversion libraries will be imported inside functions so the app
# can start even if they are not installed; we provide clear error handling
# when conversion is attempted.

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'documents'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
IMAGES_FOLDER = os.path.join('static', 'uploads')
os.makedirs(IMAGES_FOLDER, exist_ok=True)

# Toggle whether to set Word's limited highlight index (WD_COLOR_INDEX).
# When False we rely only on run shading (w:shd) to preserve exact hex colors
# which avoids several different colors collapsing to the same highlight index.
USE_WD_HIGHLIGHT = False


# --- Helpers for conversions that preserve highlights ---
def _get_closest_color(r, g, b):
    """Encuentra el color más cercano en la paleta de WD_COLOR_INDEX"""
    from docx.enum.text import WD_COLOR_INDEX
    
    # Define los colores RGB para cada WD_COLOR_INDEX mapeando a los colores de TinyMCE
    # Usando solo los colores disponibles en WD_COLOR_INDEX
    color_map = {
        # Colores claros y medios
        WD_COLOR_INDEX.YELLOW: (251, 238, 184),      # Light yellow
        WD_COLOR_INDEX.BRIGHT_GREEN: (191, 237, 210), # Light green
        WD_COLOR_INDEX.TURQUOISE: (194, 224, 244),   # Light blue
        WD_COLOR_INDEX.PINK: (248, 202, 198),        # Light red
        
        # Colores medios
        WD_COLOR_INDEX.GREEN: (45, 194, 107),        # Green
        WD_COLOR_INDEX.RED: (224, 62, 45),           # Red
        WD_COLOR_INDEX.BLUE: (53, 152, 219),         # Blue
        WD_COLOR_INDEX.TEAL: (22, 145, 121),         # Turquoise
        
        # Grises
        WD_COLOR_INDEX.GRAY_25: (236, 240, 241),     # Light gray
        WD_COLOR_INDEX.GRAY_50: (126, 140, 141),     # Dark gray
    }
    
    # Encuentra el color más cercano
    min_distance = float('inf')
    closest_color = WD_COLOR_INDEX.YELLOW  # Default
    
    for wd_color, (cr, cg, cb) in color_map.items():
        # Calcula distancia euclidiana en espacio RGB
        distance = ((r - cr) ** 2 + (g - cg) ** 2 + (b - cb) ** 2) ** 0.5
        if distance < min_distance:
            min_distance = distance
            closest_color = wd_color
            
    return closest_color

def _normalize_color(color_str):
    """Normaliza un color a formato hex y encuentra el WD_COLOR_INDEX más cercano"""
    if not color_str:
        return None, None
        
    color_str = color_str.lower().strip()
    r, g, b = 255, 255, 0  # Default amarillo
    
    # Procesar valor RGB
    if color_str.startswith('rgb'):
        try:
            parts = color_str.strip('rgb()').split(',')
            r, g, b = map(lambda x: int(x.strip()), parts)
            hex_color = '#{:02x}{:02x}{:02x}'.format(r, g, b)
            print(f"RGB {color_str} -> {hex_color}")
            return hex_color, _get_closest_color(r, g, b)
        except Exception as e:
            print(f"Error parsing RGB: {e}")
            
    # Procesar valor hex
    elif color_str.startswith('#'):
        try:
            if len(color_str) == 4:  # formato corto #fff
                color_str = '#' + ''.join(c*2 for c in color_str[1:])
            r = int(color_str[1:3], 16)
            g = int(color_str[3:5], 16)
            b = int(color_str[5:7], 16)
            return color_str, _get_closest_color(r, g, b)
        except Exception as e:
            print(f"Error parsing hex: {e}")
            
    # Nombres de colores básicos
    elif color_str in _colors:
        hex_color = '#' + _colors[color_str]
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return hex_color, _get_closest_color(r, g, b)
        except Exception as e:
            print(f"Error parsing named color: {e}")
            
    return None, None

def _parse_background_color(style_str):
    """Parse background color from style attribute or data-mce-style"""
    if not style_str:
        return None

    # Debug del estilo completo
    print(f"Parsing style: {style_str}")
        
    style_str = style_str.lower()
    parts = [p.strip() for p in style_str.split(';') if p.strip()]
    
    for p in parts:
        if 'background-color:' in p:
            try:
                val = p.split(':', 1)[1].strip()
                print(f"Found background-color value: {val}")
                return val
                    
            except Exception as e:
                print(f"Error parsing background-color: {e}")
                continue
                
    # Si no encontramos background-color, buscar otros indicadores
    if 'highlight' in style_str:
        print("Found highlight class/attribute")
        # Por defecto usar amarillo
        return 'rgb(255,255,0)'
        
    return None


def _get_style_attributes(node):
    """Extract all style attributes from a node"""
    style = node.get('style', '')
    print(f"\n>>> Extracting styles from: {style}")
    attrs = {}
    
    if not style:
        print("No style attribute found")
        return attrs
        
    parts = [p.strip() for p in style.split(';') if p.strip()]
    print(f"Style parts: {parts}")
    
    for part in parts:
        if ':' not in part:
            continue
        key, val = part.split(':', 1)
        key = key.strip().lower()
        val = val.strip()  # No convertir a lower() para preservar RGB
        attrs[key] = val
        print(f"Added style: {key} = {val}")
        
    return attrs

def _map_hex_to_wd_color(hexcode):
    """Map hex colors to WD_COLOR_INDEX"""
    try:
        from docx.enum.text import WD_COLOR_INDEX
    except Exception:
        return None
        
    if not hexcode:
        return None
        
    h = hexcode.strip().lower()
    
    # Mapa de colores RGB a WD_COLOR_INDEX actualizado para TinyMCE
    mapping = {
        # Colores claros
        'rgb(191,237,210)': WD_COLOR_INDEX.BRIGHT_GREEN,  # Light green
        'rgb(251,238,184)': WD_COLOR_INDEX.YELLOW,        # Light yellow
        'rgb(248,202,198)': WD_COLOR_INDEX.PINK,          # Light red
        'rgb(194,224,244)': WD_COLOR_INDEX.TURQUOISE,     # Light blue
        'rgb(236,240,241)': WD_COLOR_INDEX.GRAY_25,       # Light gray
        
        # Colores medios
        'rgb(45,194,107)': WD_COLOR_INDEX.GREEN,          # Green
        'rgb(241,196,15)': WD_COLOR_INDEX.YELLOW,         # Yellow
        'rgb(224,62,45)': WD_COLOR_INDEX.RED,             # Red
        'rgb(185,106,217)': WD_COLOR_INDEX.TURQUOISE,     # Purple
        'rgb(53,152,219)': WD_COLOR_INDEX.BLUE,           # Blue
        'rgb(22,145,121)': WD_COLOR_INDEX.TEAL,           # Dark turquoise
        
        # Colores oscuros
        'rgb(35,111,161)': WD_COLOR_INDEX.BLUE,           # Dark blue
        'rgb(52,73,94)': WD_COLOR_INDEX.TEAL,             # Navy blue
        
        # Grises
        'rgb(206,212,217)': WD_COLOR_INDEX.GRAY_25,       # Medium gray
        'rgb(126,140,141)': WD_COLOR_INDEX.GRAY_50,       # Dark gray
    }
    
    if h in mapping:
        return mapping[h]
    
    # Para RGB en formato numérico
    if h.startswith('rgb'):
        try:
            r, g, b = map(int, h.strip('rgb()').split(','))
            rgb_str = f'rgb({r},{g},{b})'
            if rgb_str in mapping:
                return mapping[rgb_str]
        except:
            pass
            
    return None  # Si no hay coincidencia, no aplicar highlight


def _set_docx_run_shading(run, hex_color):
    """Set explicit run shading (w:shd w:fill) so Word displays exact background color.
    This is a fallback to avoid depending only on WD_COLOR_INDEX limited palette.
    """
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        # create w:shd element and set fill to RRGGBB (no #)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color.lstrip('#').upper())
        rpr = run._r.get_or_add_rPr()
        # remove existing shd if present
        for child in list(rpr):
            if child.tag.endswith('shd'):
                rpr.remove(child)
        rpr.append(shd)
    except Exception as e:
        print(f"Failed to set docx shading for {hex_color}: {e}")


def _extract_and_save_inline_images(html, upload_folder='static/uploads'):
    """Busca <img src="data:...base64..."> en el HTML, guarda cada imagen
    en disk bajo `upload_folder` y reemplaza el src por la ruta estática.

    Retorna el HTML modificado.
    """
    import base64
    import uuid
    from bs4 import BeautifulSoup
    import re

    soup = BeautifulSoup(html, 'lxml')
    os.makedirs(upload_folder, exist_ok=True)

    for img in soup.find_all('img'):
        src = img.get('src', '')
        if not src:
            continue
        # detect data URI
        if src.startswith('data:'):
            m = re.match(r'data:(image/[^;]+);base64,(.*)', src, re.S)
            if not m:
                continue
            mime = m.group(1)
            data = m.group(2)
            # choose extension from mime
            ext = 'png'
            if mime == 'image/jpeg' or mime == 'image/jpg':
                ext = 'jpg'
            elif mime == 'image/gif':
                ext = 'gif'
            elif mime == 'image/svg+xml':
                ext = 'svg'

            filename = f"{uuid.uuid4().hex}.{ext}"
            filepath = os.path.join(upload_folder, filename)
            try:
                with open(filepath, 'wb') as fh:
                    fh.write(base64.b64decode(data))
                # set the img src to a relative static path so the saved HTML
                # works both when served by Flask and when opened directly
                # from disk. From `documents/` the correct relative path to
                # the `static/uploads` folder is `../static/uploads/...`.
                api_url = None
                try:
                    api_url = url_for('serve_image', filename=filename)
                    img['src'] = api_url
                    img['data-local-src'] = f"../{upload_folder}/{filename}"
                    print(f"Inline image saved and src rewritten to API: {img['src']}")
                except Exception:
                    # fallback to relative static path if url_for isn't available
                    rel = f"../{upload_folder}/{filename}"
                    img['src'] = rel
                    img['data-local-src'] = rel
                    print(f"Inline image saved and src set to relative path: {img['src']}")
                print(f"Saved inline image to {filepath}")
            except Exception as e:
                print(f"Failed to save inline image: {e}")

    return str(soup)

def _convert_local_srcs_to_api(html):
    """Rewrite local image srcs that point to the uploads folder to use
    the `/api/images/<filename>` endpoint so they behave like remote URLs.
    Handles paths like `../static/uploads/<file>`, `/static/uploads/<file>`,
    `../api/images/<file>`, `/api/images/<file>`, and `file://.../static/uploads/<file>`.
    """
    try:
        from bs4 import BeautifulSoup
    except Exception:
        return html

    soup = BeautifulSoup(html, 'lxml')
    for img in soup.find_all('img'):
        src = img.get('src', '')
        if not src:
            continue
        # find filename after static/uploads/
        m = re.search(r'static/uploads/([^"\'>]+)', src)
        if m:
            filename = m.group(1)
            try:
                api_url = url_for('serve_image', filename=filename)
                img['src'] = api_url
                img['data-local-src'] = f"../static/uploads/{filename}"
                print(f"Rewrote local src {src} -> {img['src']}")
            except Exception:
                img['src'] = f"/api/images/{filename}"
                img['data-local-src'] = f"../static/uploads/{filename}"
                print(f"Rewrote local src {src} -> {img['src']} (fallback)")
            continue

        # find filename after api/images/ (may be relative like ../api/images/...)
        m2 = re.search(r'api/images/([^"\'>\\]+)', src)
        if m2:
            filename = m2.group(1)
            try:
                api_url = url_for('serve_image', filename=filename)
                img['src'] = api_url
                img['data-local-src'] = f"../static/uploads/{filename}"
                print(f"Rewrote api/images src {src} -> {img['src']}")
            except Exception:
                img['src'] = f"/api/images/{filename}"
                img['data-local-src'] = f"../static/uploads/{filename}"
                print(f"Rewrote api/images src {src} -> {img['src']} (fallback)")
            continue

    return str(soup)

@app.route('/api/images/<path:filename>')
def serve_image(filename):
    """Serve images stored in `static/uploads` via a simple API route.

    Using `send_from_directory` allows us to keep images outside the
    application's document tree and reference them via `/api/images/...`.
    """
    # Security: prevent path traversal
    filename = os.path.normpath(filename)
    if filename.startswith('..'):
        return "Invalid filename", 400
    return send_from_directory(IMAGES_FOLDER, filename)
def _resolve_image_src_to_path(src, base_dir=None):
    """Resolve an <img> src to a filesystem path.

    - If `src` is an absolute web-root path like `/static/uploads/...`, resolve
      it relative to the project root (`os.getcwd()`).
    - If `src` is a file:// URL, strip the scheme and return the path.
    - Otherwise treat `src` as relative to `base_dir` (if provided) or the
      current working directory.
    Returns a normalized filesystem path (may not exist).
    """
    if not src:
        return None
    src = src.strip()
    # file:// URL
    if src.startswith('file://'):
        path = src[len('file://'):]
        return os.path.normpath(path)
    # http(s) - try to download into IMAGES_FOLDER so we can embed it
    if src.startswith('http://') or src.startswith('https://'):
        try:
            import urllib.request
            import mimetypes
            # fetch bytes
            with urllib.request.urlopen(src) as resp:
                data = resp.read()
                info = resp.info()
                ctype = info.get_content_type() if hasattr(info, 'get_content_type') else None

            # choose extension
            ext = None
            if ctype:
                ext = mimetypes.guess_extension(ctype)
            if not ext:
                # try imghdr
                try:
                    import imghdr
                    ext = imghdr.what(None, h=data)
                    if ext:
                        ext = '.' + ext
                except Exception:
                    ext = '.bin'

            if not os.path.exists(IMAGES_FOLDER):
                os.makedirs(IMAGES_FOLDER, exist_ok=True)
            fname = f"{uuid.uuid4().hex}{ext}"
            fpath = os.path.join(IMAGES_FOLDER, fname)
            with open(fpath, 'wb') as fh:
                fh.write(data)
            return os.path.normpath(fpath)
        except Exception as e:
            print(f"Failed to download remote image {src}: {e}")
            return None
    # absolute web-root path: /static/...
    # If the src contains our API route (e.g. '/api/images/...' or '../api/images/...')
    m_api = re.search(r'api/images/([^"\'>\\]+)', src)
    if m_api:
        filename = m_api.group(1)
        return os.path.normpath(os.path.join(os.getcwd(), IMAGES_FOLDER, filename))

    if src.startswith('/'):
        # Special-case our API route handled above; otherwise treat as web-root path
        return os.path.normpath(os.path.join(os.getcwd(), src.lstrip('/')))
    # relative path - resolve against base_dir if provided
    base = base_dir or os.getcwd()
    return os.path.normpath(os.path.join(base, src))


def _download_remote_images_in_html(html, upload_folder='static/uploads'):
    """Find <img src="http(s)..."> in HTML, download them into `upload_folder`
    and replace src with a relative path `../static/uploads/<file>` so the
    saved HTML references the local copy.
    """
    try:
        from bs4 import BeautifulSoup
        import urllib.request, mimetypes, imghdr
        import uuid
    except Exception:
        return html

    soup = BeautifulSoup(html, 'lxml')
    os.makedirs(upload_folder, exist_ok=True)
    changed = False
    for img in soup.find_all('img'):
        src = img.get('src', '').strip()
        if not src:
            continue
        if src.startswith('http://') or src.startswith('https://'):
            try:
                with urllib.request.urlopen(src) as resp:
                    data = resp.read()
                    info = resp.info()
                    ctype = info.get_content_type() if hasattr(info, 'get_content_type') else None

                ext = None
                if ctype:
                    ext = mimetypes.guess_extension(ctype)
                if not ext:
                    try:
                        t = imghdr.what(None, h=data)
                        if t:
                            ext = '.' + t
                    except Exception:
                        ext = '.bin'

                fname = f"{uuid.uuid4().hex}{ext or '.bin'}"
                fpath = os.path.join(upload_folder, fname)
                with open(fpath, 'wb') as fh:
                    fh.write(data)
                img['src'] = f"../{upload_folder}/{fname}"
                changed = True
                print(f"Downloaded remote image {src} -> {fpath}")
            except Exception as e:
                print(f"Failed to download embedded remote image {src}: {e}")
                continue

    return str(soup)


def _get_image_size(path):
    """Return (width, height) in pixels for common image types.

    Tries Pillow first, then falls back to simple header parsing for PNG, JPEG and GIF.
    Returns (None, None) if dimensions cannot be determined.
    """
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.width, im.height
    except Exception:
        pass

    try:
        with open(path, 'rb') as fh:
            data = fh.read(64*1024)
        # PNG
        if data[:8] == b'\x89PNG\r\n\x1a\n':
            import struct
            # IHDR chunk at offset 8..25, width at 16..19, height at 20..23
            width = struct.unpack('>I', data[16:20])[0]
            height = struct.unpack('>I', data[20:24])[0]
            return width, height
        # JPEG - parse markers to find SOF0/2
        if data[:2] == b'\xff\xd8':
            import io, struct
            f = io.BytesIO(data)
            f.read(2)
            while True:
                marker_bytes = f.read(2)
                if len(marker_bytes) < 2:
                    break
                marker, = struct.unpack('>H', marker_bytes)
                # Skip padding
                while marker == 0xFFFF:
                    marker, = struct.unpack('>H', f.read(2))
                # SOF markers range
                if 0xFFC0 <= marker <= 0xFFC3 or 0xFFC5 <= marker <= 0xFFC7 or 0xFFC9 <= marker <= 0xFFCB or 0xFFCD <= marker <= 0xFFCF:
                    length_bytes = f.read(2)
                    if len(length_bytes) < 2:
                        break
                    length = struct.unpack('>H', length_bytes)[0]
                    precision = f.read(1)
                    h_bytes = f.read(2)
                    w_bytes = f.read(2)
                    if len(h_bytes) == 2 and len(w_bytes) == 2:
                        height = struct.unpack('>H', h_bytes)[0]
                        width = struct.unpack('>H', w_bytes)[0]
                        return width, height
                    break
                else:
                    # skip this segment
                    length_bytes = f.read(2)
                    if len(length_bytes) < 2:
                        break
                    length = struct.unpack('>H', length_bytes)[0]
                    f.read(length-2)
        # GIF
        if data[:6] in (b'GIF87a', b'GIF89a'):
            import struct
            width = struct.unpack('<H', data[6:8])[0]
            height = struct.unpack('<H', data[8:10])[0]
            return width, height
    except Exception:
        pass

    return None, None


def _ensure_file_view_script(html):
    """Inject a small fallback script into saved HTML so that when the
    file is opened via file:// the images with `data-local-src` are used.

    The function is idempotent: it won't inject the script more than once.
    """
    try:
        if 'data-local-src' not in html:
            return html
        marker = '<!-- FILE_VIEW_FALLBACK_SCRIPT -->'
        if marker in html:
            return html
        script = (
            "\n<!-- FILE_VIEW_FALLBACK_SCRIPT -->\n"
            "<script>\n"
            "(function(){\n"
            "  try{\n"
            "    if (location.protocol === 'file:'){\n"
            "      Array.prototype.forEach.call(document.querySelectorAll('img[data-local-src]'), function(img){\n"
            "        var local = img.getAttribute('data-local-src'); if(local) img.src = local; });\n"
            "    }\n"
            "  }catch(e){}\n"
            "})();\n"
            "</script>\n"
        )
        import re
        # Insert before </body> if present (case-insensitive), otherwise append.
        if re.search(r'</body>', html, flags=re.IGNORECASE):
            html = re.sub(r'</body>', script + '</body>', html, flags=re.IGNORECASE)
        else:
            html = html + script
        return html
    except Exception:
        return html


def html_to_docx(html, out_path, base_dir=None):
    """Convert HTML to DOCX preserving highlights and formatting"""
    from bs4 import BeautifulSoup
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Pt, RGBColor, Cm

    soup = BeautifulSoup(html, 'lxml')
    doc = Document()
    # Reduce default margins (python-docx default is ~2.54cm / 1in).
    try:
        section = doc.sections[0]
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
    except Exception as e:
        print(f"Could not set DOCX margins: {e}")

    def apply_text_formatting(run, node):
        """Apply text formatting (bold, italic, etc) to a run"""
        # Obtener todos los estilos
        styles = _get_style_attributes(node)
        
        # Procesar estilo font-weight para negrita
        if 'font-weight' in styles:
            run.bold = styles['font-weight'] in ('bold', '700')
            
        # Procesar estilo text-decoration para subrayado
        if 'text-decoration' in styles:
            run.underline = 'underline' in styles['text-decoration'].lower()
            
        # Procesar estilo font-style para cursiva
        if 'font-style' in styles:
            run.italic = styles['font-style'] == 'italic'
            
        # Procesar background-color para highlights
        if 'background-color' in styles:
            print(f"Processing background-color: {styles['background-color']}")  # Debug
            hex_color, highlight = _normalize_color(styles['background-color'])
            print(f"Normalized to hex: {hex_color}, highlight: {highlight}")  # Debug
            if hex_color:
                # Try to set Word's limited highlight index (for compatibility)
                if USE_WD_HIGHLIGHT and highlight:
                    try:
                        run.font.highlight_color = highlight
                        print(f"Setting highlight color to: {highlight}")
                    except Exception as e:
                        print(f"Could not set WD highlight: {e}")
                # Also set precise shading (w:shd) so Word displays exact background
                _set_docx_run_shading(run, hex_color)

    def process_text_node(node, parent_paragraph, base_dir=None):
        """Process text nodes preserving formatting"""
        if not node:
            return
            
        # Debug: imprimir información del nodo
        if node.name:
            print(f"\n=== Processing node: {node.name} ===")
            print(f"Text content: {node.get_text()}")
            print(f"Attributes: {node.attrs}")
            print(f"Style: {node.get('style', '')}")
            print(f"Data-mce-style: {node.get('data-mce-style', '')}")
            
        # Nodos de texto directo (preservar espacios)
        if not node.name:
            text = str(node)  # No usar strip() para preservar espacios
            if text:  # Incluir incluso si son solo espacios
                parent_paragraph.add_run(text)
            return
                
        # Procesar imágenes embebidas
        if node.name == 'img':
            src = node.get('src', '')
            print(f"DOCX: encountered <img> with src='{src}'")
            if src:
                file_path = _resolve_image_src_to_path(src, base_dir=base_dir)
                print(f"DOCX: resolved image path: {file_path}")
                if file_path is None:
                    print(f"Skipping remote image for DOCX: {src}")
                    return
                try:
                    # Intentar insertar la imagen en el párrafo (inline)
                    try:
                        run = parent_paragraph.add_run()
                        run.add_picture(file_path)
                    except Exception:
                        # Fallback: añadir imagen como párrafo independiente
                        print(f"DOCX: fallback add_picture for {file_path}")
                        doc.add_picture(file_path)
                except Exception as e:
                    print(f"Failed to add image to DOCX from {file_path}: {e}")
            return

        

        # Procesar elementos con formato
        if node.name in ('span', 'mark', 'strong', 'b', 'em', 'i', 'u'):
            text = node.get_text()  # No usar strip() para preservar espacios
            if text:  # Incluir incluso si son solo espacios
                run = parent_paragraph.add_run(text)
                print(f"\n--- Formatting run for text: {text} ---")
                
                # Obtener todos los estilos disponibles
                style = node.get('style', '')
                mce_style = node.get('data-mce-style', '')
                print(f"Style attribute: {style}")
                print(f"MCE style attribute: {mce_style}")
                
                # Aplicar formato basado en etiquetas HTML
                if node.name in ('strong', 'b'):
                    run.bold = True
                if node.name in ('em', 'i'):
                    run.italic = True
                if node.name == 'u':
                    run.underline = True
                
                # Extraer y procesar estilos
                styles = _get_style_attributes(node)
                print(f"Extracted styles: {styles}")
                
                # Procesar background-color
                if 'background-color' in styles:
                    color_value = styles['background-color']
                    print(f"Found background-color: {color_value}")
                    hex_color, highlight = _normalize_color(color_value)
                    print(f"Normalized to hex: {hex_color}, highlight: {highlight}")
                    if hex_color:
                        if USE_WD_HIGHLIGHT and highlight:
                            try:
                                run.font.highlight_color = highlight
                                print(f"Setting highlight color to: {highlight}")
                            except Exception as e:
                                print(f"Could not set WD highlight: {e}")
                        # set exact shading too
                        _set_docx_run_shading(run, hex_color)
                    else:
                        print("Warning: Failed to normalize color for DOCX shading")
                
                # Si no se aplicó highlight, intentar con data-mce-style
                if not run.font.highlight_color and mce_style:
                    print(f"Trying data-mce-style: {mce_style}")
                    node_with_mce = type('Node', (), {'get': lambda s, x: mce_style if x == 'style' else None})()
                    apply_text_formatting(run, node_with_mce)
            return
        
        # Procesar imágenes en elementos inline dentro de bloques (por si aparecen como hijos)
        if node.name == 'img':
            src = node.get('src', '')
            if src:
                file_path = _resolve_image_src_to_path(src, base_dir=base_dir)
                if file_path is None:
                    print(f"Skipping remote nested image for DOCX: {src}")
                    return
                try:
                    try:
                        run = parent_paragraph.add_run()
                        run.add_picture(file_path)
                    except Exception:
                        doc.add_picture(file_path)
                except Exception as e:
                    print(f"Failed to add nested image to DOCX from {file_path}: {e}")
            return
            
        # Procesar otros elementos de forma recursiva
        for child in node.children:
            process_text_node(child, parent_paragraph, base_dir=base_dir)

    # Procesar bloques de texto
    for el in soup.find_all(['h1', 'h2', 'h3', 'p', 'li', 'div']):
        p = doc.add_paragraph()
        
        # Procesar cada nodo hijo preservando formato
        for node in el.children:
            if node.name == 'br':
                p.add_run('\n')
            else:
                process_text_node(node, p, base_dir=base_dir)
        
        # Nota: no añadimos un párrafo vacío extra aquí; cada elemento bloque
        # ya creó su propio párrafo `p` (evita líneas en blanco adicionales).
            
    # Save the document
    doc.save(out_path)
def html_to_odt(html, out_path, base_dir=None):
    """Convert HTML to ODT preserving highlights and formatting"""
    from bs4 import BeautifulSoup
    from odf.opendocument import OpenDocumentText
    from odf.text import P, Span
    from odf.style import Style, TextProperties

    soup = BeautifulSoup(html, 'lxml')
    doc = OpenDocumentText()

    # Diccionario de estilos base
    base_styles = {
        'bold': {'name': 'Bold', 'family': 'text', 'props': {'fontweight': 'bold'}},
        'italic': {'name': 'Italic', 'family': 'text', 'props': {'fontstyle': 'italic'}},
        'underline': {'name': 'Underline', 'family': 'text', 'props': {'textunderlinestyle': 'solid', 'textunderlinewidth': 'auto'}},
    }
    # Creamos solo los estilos base; los highlights se crearán dinámicamente
    for style_info in base_styles.values():
        style = Style(name=style_info['name'], family=style_info['family'])
        style.addElement(TextProperties(**style_info['props']))
        doc.automaticstyles.addElement(style)

    def get_style_name(node):
        """Determinar el nombre de estilo basado en el formato. Crea estilos de highlight dinámicos."""
        styles = _get_style_attributes(node)
        formats = []

        # Detectar formatos básicos
        if node.name in ('strong', 'b') or styles.get('font-weight') in ('bold', '700'):
            formats.append('bold')
        if node.name in ('em', 'i') or styles.get('font-style') == 'italic':
            formats.append('italic')
        if node.name == 'u' or 'underline' in styles.get('text-decoration', ''):
            formats.append('underline')

        # Detectar background-color y crear estilo dinámico
        bg = styles.get('background-color') or node.get('data-mce-style', '')
        if bg:
            # Si la cadena ya es un 'rgb(...)' o '#hex', úsala tal cual.
            # Sólo usar _parse_background_color cuando la cadena contiene el prefijo 'background-color:'
            if 'background-color' in bg:
                bg_val = _parse_background_color(bg)
            else:
                bg_val = bg
            if bg_val:
                try:
                    hex_color, _ = _normalize_color(bg_val)
                    # Si normalize no devolvió hex, intentar conversión directa
                    if not hex_color and bg_val.startswith('rgb'):
                        parts = bg_val.strip('rgb()').split(',')
                        r, g, b = map(lambda x: int(x.strip()), parts)
                        hex_color = '#{:02x}{:02x}{:02x}'.format(r, g, b)

                    if hex_color:
                        style_name = f'highlight-{hex_color[1:]}'
                        existing = [s.getAttribute('name') for s in doc.automaticstyles.childNodes]
                        if style_name not in existing:
                            st = Style(name=style_name, family='text')
                            # Use backgroundcolor kwarg which odfpy maps correctly to the
                            # text-properties fo:background-color attribute in the ODT XML.
                            tp = TextProperties(backgroundcolor=hex_color)
                            st.addElement(tp)
                            doc.automaticstyles.addElement(st)
                            print(f"Created new style: {style_name}")
                            # Debug: listar estilos automáticos tras creación
                            existing_after = [s.getAttribute('name') for s in doc.automaticstyles.childNodes]
                            print(f"Automatic styles after adding: {existing_after}")
                        # Crear combinación con formatos si es necesario
                        if formats:
                            combined = f"{style_name}-{'-'.join(formats)}"
                            if combined not in existing:
                                props = {}
                                if 'bold' in formats:
                                    props['fontweight'] = 'bold'
                                if 'italic' in formats:
                                    props['fontstyle'] = 'italic'
                                if 'underline' in formats:
                                    props['textunderlinestyle'] = 'solid'
                                    props['textunderlinewidth'] = 'auto'
                                st2 = Style(name=combined, family='text')
                                # Create TextProperties with both formatting props and background
                                tp2 = TextProperties(backgroundcolor=hex_color, **props)
                                st2.addElement(tp2)
                                doc.automaticstyles.addElement(st2)
                                print(f"Created combined style: {combined}")
                                return combined
                        return style_name
                except Exception as e:
                    print(f"Error creando estilo ODT para color {bg_val}: {e}")

        if formats:
            return formats[0].title()
        return None

    def process_text_node(node, parent_paragraph, base_dir=None):
        """Process text nodes preserving formatting"""
        if not node:
            return
            
        # Nodos de texto directo (preservar espacios)
        if not node.name:
            text = str(node)  # No usar strip() para preservar espacios
            if text:
                parent_paragraph.addText(text)
            return
                
        # Procesar imágenes embebidas en ODT
        if node.name == 'img':
            src = node.get('src', '')
            print(f"ODT: encountered <img> with src='{src}'")
            if src:
                file_path = _resolve_image_src_to_path(src, base_dir=base_dir)
                print(f"ODT: resolved image path: {file_path}")
                if file_path is None:
                    print(f"Skipping remote image for ODT: {src}")
                    return
                try:
                    from odf.draw import Frame, Image
                    import mimetypes
                    with open(file_path, 'rb') as fimg:
                        content = fimg.read()
                    mediatype, _ = mimetypes.guess_type(file_path)
                    if mediatype is None:
                        mediatype = 'application/octet-stream'
                    try:
                        manifestfn = doc.addPictureFromString(content, mediatype)
                        href = manifestfn
                    except Exception as e:
                        print(f"addPictureFromString failed for ODT: {e}, trying addPictureFromFile")
                        try:
                            manifestfn = doc.addPictureFromFile(file_path)
                            href = manifestfn
                        except Exception as e2:
                            print(f"addPictureFromFile also failed: {e2}")
                            href = file_path

                    # Determine image pixel size to preserve aspect ratio
                    w_px, h_px = _get_image_size(file_path)
                    if w_px and h_px:
                        # assume 96 dpi for conversion to cm
                        dpi = 96.0
                        width_cm = (w_px / dpi) * 2.54
                        height_cm = (h_px / dpi) * 2.54
                        max_width_cm = 15.0
                        if width_cm > max_width_cm:
                            scale = max_width_cm / width_cm
                            width_cm *= scale
                            height_cm *= scale
                        frame = Frame(width=f"{width_cm:.2f}cm", height=f"{height_cm:.2f}cm")
                    else:
                        frame = Frame(width="6cm", height="4cm")

                    image = Image(href=href)
                    frame.addElement(image)
                    parent_paragraph.addElement(frame)
                except Exception as e:
                    print(f"Failed to add image to ODT from {file_path}: {e}")
            return

        # Procesar elementos con formato
        if node.name in ('span', 'mark', 'strong', 'b', 'em', 'i', 'u'):
            text = node.get_text()  # No usar strip() para preservar espacios
            if text:
                style_name = get_style_name(node)
                print(f"Applying ODT style: {style_name}")
                existing_styles = [s.getAttribute('name') for s in doc.automaticstyles.childNodes]
                print(f"Existing automatic styles before span: {existing_styles}")
                if style_name:
                    span = Span(stylename=style_name)
                    # asegurar que el atributo text:style-name está presente
                    try:
                        span.setAttribute('text:style-name', style_name)
                    except Exception:
                        pass
                    span.addText(text)
                    parent_paragraph.addElement(span)
                else:
                    parent_paragraph.addText(text)
            return
            
        # Para otros elementos inline, procesar su contenido
        for child in node.children:
            process_text_node(child, parent_paragraph, base_dir=base_dir)

    # Procesar bloques de texto
    for el in soup.find_all(['h1', 'h2', 'h3', 'p', 'li', 'div']):
        p = P()
        
        # Procesar cada nodo hijo preservando formato
        for node in el.children:
            if node.name == 'br':
                p.addText('\n')
            else:
                process_text_node(node, p, base_dir=base_dir)

        # Añadir párrafo al documento
        doc.text.addElement(p)
        
        # Salto adicional después de bloques (excepto items de lista)
        if el.name != 'li':
            doc.text.addElement(P())
            
    # Save the document
    doc.save(out_path)

# --- end helpers ---

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/documents')
def list_documents():
    files = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.html')]
    return render_template('list.html', files=files)

@app.route('/explorer')
def explorer():
    files = os.listdir(UPLOAD_FOLDER)
    # Build list of document base names from existing HTML files so we can
    # show all available export options (including PDF) for each document.
    bases = set()
    for f in files:
        if f.endswith('.html'):
            bases.add(f.rsplit('.', 1)[0])

    grouped = {}
    for base in bases:
        formats = []
        # html always available for bases found
        formats.append('html')
        # include odt/docx only if they exist
        if os.path.exists(os.path.join(UPLOAD_FOLDER, base + '.odt')):
            formats.append('odt')
        if os.path.exists(os.path.join(UPLOAD_FOLDER, base + '.docx')):
            formats.append('docx')
        # always offer pdf (we will generate on-demand if it's missing)
        formats.append('pdf')
        grouped[base] = formats

    return render_template('explorer.html', grouped=grouped)

@app.route('/edit/<filename>')
def edit(filename):
    if not filename.endswith('.html'):
        return "Solo se pueden editar archivos HTML", 400
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    with open(filepath, 'r') as f:
        content = f.read()
    # Fix any absolute/static-root image paths saved previously so the
    # editor shows images correctly even when the HTML file is opened
    # directly (file://) or when served by Flask. Convert:
    #  - "/static/uploads/..." -> "../static/uploads/..."
    #  - "file:///C:/.../static/uploads/..." -> "../static/uploads/..."
    try:
        import re
        # replace leading /static/uploads/ occurrences
        content = re.sub(r'src=["\']?/static/uploads/', 'src="../static/uploads/', content)
        # replace file:// style absolute paths that include static/uploads
        content = re.sub(r'src=["\']?file://[^"\']*?/static/uploads/', 'src="../static/uploads/', content)
    except Exception:
        pass

    # Also convert local upload references to API URLs so the editor loads them
    try:
        content = _convert_local_srcs_to_api(content)
    except Exception:
        pass

    return render_template('form.html', content=content, filename=filename)

@app.route('/generate', methods=['POST'])
def generate():
    content = request.form['content']
    # Debug: imprimir el HTML recibido para ver la estructura
    print("HTML recibido:", content)
    
    # Asegurar que tenemos HTML bien formado con los highlights
    if not content.strip().startswith('<'):
        content = f'<div>{content}</div>'
    elif not any(tag in content.lower() for tag in ['<body', '<div', '<p']):
        content = f'<div>{content}</div>'
    
    docname = request.form['docname'].strip().replace(' ', '_')
    filename = secure_filename(request.form.get('filename') or f"{docname}.html")

    html_path = os.path.join(UPLOAD_FOLDER, filename)
    odt_path = html_path.replace('.html', '.odt')
    docx_path = html_path.replace('.html', '.docx')

    # Extract any inline (data URI) images pasted from the editor and save
    # them to the static uploads folder, replacing the <img> src attributes
    # with the corresponding `/static/uploads/...` paths.
    try:
        # Only extract inline data: images pasted from the editor. We no
        # longer download remote http(s) images into the saved HTML; remote
        # image URLs remain external. Converters that need to embed remote
        # images will download them on-demand during conversion via
        # `_resolve_image_src_to_path`.
        content = _extract_and_save_inline_images(content)
        # Ensure any local references to uploads are rewritten to the API
        content = _convert_local_srcs_to_api(content)
        # Inject a small fallback script so static HTML opened via file://
        # will swap `src` to `data-local-src` and display images locally.
        content = _ensure_file_view_script(content)
    except Exception as e:
        print(f"Warning: failed to extract inline images: {e}")

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(content)

    messages = []

    # DOCX conversion using python-docx with highlight preservation
    try:
        with open(html_path, 'r', encoding='utf-8') as fh:
            html = fh.read()
        # Prepare a copy of the HTML for conversion where we remove any
        # injected file:// fallback scripts. Those scripts are useful for
        # static viewing but must not be embedded as text in DOCX/ODT.
        try:
            try:
                from bs4 import BeautifulSoup
                soup_for_conv = BeautifulSoup(html, 'lxml')
            except Exception:
                from bs4 import BeautifulSoup
                soup_for_conv = BeautifulSoup(html, 'html.parser')

            # Remove script tags that implement the file:// fallback
            removed = False
            for script in list(soup_for_conv.find_all('script')):
                content = script.string or ''
                if 'location.protocol' in content or 'FILE_VIEW_FALLBACK_SCRIPT' in str(script.previous_sibling):
                    script.decompose()
                    removed = True

            if removed:
                print('Removed file-view fallback script from HTML before DOCX conversion')

            html_for_docx = str(soup_for_conv)
        except Exception:
            html_for_docx = html

        try:
            html_to_docx(html_for_docx, docx_path, base_dir=os.path.dirname(html_path))
            messages.append('DOCX generado (preservando highlights)')
        except Exception as e:
            messages.append('No se pudo generar DOCX: ' + str(e))
    except Exception as e:
        messages.append('Error leyendo HTML para DOCX: ' + str(e))

    # ODT conversion using odfpy + BeautifulSoup with highlight styles
    try:
        with open(html_path, 'r', encoding='utf-8') as fh:
            html = fh.read()
        # Remove the fallback script before ODT conversion as well
        try:
            try:
                from bs4 import BeautifulSoup
                soup_for_conv = BeautifulSoup(html, 'lxml')
            except Exception:
                from bs4 import BeautifulSoup
                soup_for_conv = BeautifulSoup(html, 'html.parser')

            removed = False
            for script in list(soup_for_conv.find_all('script')):
                content = script.string or ''
                if 'location.protocol' in content or 'FILE_VIEW_FALLBACK_SCRIPT' in str(script.previous_sibling):
                    script.decompose()
                    removed = True
            if removed:
                print('Removed file-view fallback script from HTML before ODT conversion')

            html_for_odt = str(soup_for_conv)
        except Exception:
            html_for_odt = html

        try:
            html_to_odt(html_for_odt, odt_path, base_dir=os.path.dirname(html_path))
            messages.append('ODT generado (preservando highlights)')
        except Exception as e:
            messages.append('No se pudo generar ODT: ' + str(e))
    except Exception as e:
        messages.append('Error leyendo HTML para ODT: ' + str(e))

    # PDF conversion using PyMuPDF (fitz)
    try:
        import fitz
        from bs4 import BeautifulSoup
        pdf_path = html_path.replace('.html', '.pdf')
        
        # Parse HTML to extract text and structure
        with open(html_path, 'r', encoding='utf-8') as fh:
            soup = BeautifulSoup(fh.read(), 'lxml')
        
        # Create PDF
        doc = fitz.open()
        page = doc.new_page()
        
        # Extract text maintaining basic structure
        text_blocks = []
        for element in soup.find_all(['h1', 'h2', 'h3', 'p', 'li']):
            # Add spacing before headers
            if element.name.startswith('h'):
                text_blocks.append('')
            
            text = element.get_text().strip()
            if text:
                text_blocks.append(text)
                # Add spacing after each block
                text_blocks.append('')
        
        # Join all text with newlines and insert into PDF
        content = '\n'.join(text_blocks)
        page.insert_text((50, 50), content, fontname="helv", fontsize=11)
        
        try:
            doc.save(pdf_path)
            messages.append('PDF generado')
        except Exception as e:
            messages.append('Error al guardar PDF: ' + str(e))
        finally:
            doc.close()
    except Exception as e:
        messages.append('PyMuPDF no disponible: ' + str(e))

    flash(' / '.join(messages))

    return render_template('form.html', content=content, filename=filename)

@app.route('/documents/<filename>')
def download(filename):
    # If a converted format is requested and doesn't exist yet, try to generate it
    target_path = os.path.join(UPLOAD_FOLDER, filename)
    base, ext = filename.rsplit('.', 1)

    if not os.path.exists(target_path) and ext in ('pdf', 'docx', 'odt'):
        html_path = os.path.join(UPLOAD_FOLDER, base + '.html')
        if not os.path.exists(html_path):
            return f"Archivo fuente HTML '{base}.html' no encontrado para generar {ext}", 404

        # Generate requested format
        try:
            if ext == 'docx':
                with open(html_path, 'r', encoding='utf-8') as fh:
                    html = fh.read()
                html_to_docx(html, target_path, base_dir=os.path.dirname(html_path))

            elif ext == 'odt':
                with open(html_path, 'r', encoding='utf-8') as fh:
                    html = fh.read()
                html_to_odt(html, target_path, base_dir=os.path.dirname(html_path))

            elif ext == 'pdf':
                import fitz
                from bs4 import BeautifulSoup
                
                # Parse HTML
                with open(html_path, 'r', encoding='utf-8') as fh:
                    soup = BeautifulSoup(fh.read(), 'lxml')
                
                # Create PDF
                doc = fitz.open()
                page = doc.new_page()
                
                # Extract text maintaining structure
                text_blocks = []
                for element in soup.find_all(['h1', 'h2', 'h3', 'p', 'li']):
                    if element.name.startswith('h'):
                        text_blocks.append('')
                    text = element.get_text().strip()
                    if text:
                        text_blocks.append(text)
                        text_blocks.append('')
                
                # Insert text into PDF
                content = '\n'.join(text_blocks)
                try:
                    page.insert_text((50, 50), content, fontname="helv", fontsize=11)
                    doc.save(target_path)
                except Exception as e:
                    return f"Error al generar PDF: {e}", 500
                finally:
                    doc.close()

        except Exception as e:
            return f"Error al generar {ext}: {e}", 500

    if not os.path.exists(target_path):
        return "Archivo no encontrado", 404

    # Serve HTML inline so the browser renders it and image URLs (/api/images/...) work
    # For other formats (docx/odt/pdf) keep attachment behavior.
    if ext == 'html':
        return send_file(target_path)
    else:
        return send_file(target_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
