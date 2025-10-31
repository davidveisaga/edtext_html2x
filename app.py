from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
import re

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


def html_to_docx(html, out_path):
    """Convert HTML to DOCX preserving highlights and formatting"""
    from bs4 import BeautifulSoup
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Pt, RGBColor

    soup = BeautifulSoup(html, 'lxml')
    doc = Document()

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

    def process_text_node(node, parent_paragraph):
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
            
        # Procesar otros elementos de forma recursiva
        for child in node.children:
            process_text_node(child, parent_paragraph)

    # Procesar bloques de texto
    for el in soup.find_all(['h1', 'h2', 'h3', 'p', 'li', 'div']):
        p = doc.add_paragraph()
        
        # Procesar cada nodo hijo preservando formato
        for node in el.children:
            if node.name == 'br':
                p.add_run('\n')
            else:
                process_text_node(node, p)
        
        # Nota: no añadimos un párrafo vacío extra aquí; cada elemento bloque
        # ya creó su propio párrafo `p` (evita líneas en blanco adicionales).
            
    # Save the document
    doc.save(out_path)
def html_to_odt(html, out_path):
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

    def process_text_node(node, parent_paragraph):
        """Process text nodes preserving formatting"""
        if not node:
            return
            
        # Nodos de texto directo (preservar espacios)
        if not node.name:
            text = str(node)  # No usar strip() para preservar espacios
            if text:
                parent_paragraph.addText(text)
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
            process_text_node(child, parent_paragraph)

    # Procesar bloques de texto
    for el in soup.find_all(['h1', 'h2', 'h3', 'p', 'li', 'div']):
        p = P()
        
        # Procesar cada nodo hijo preservando formato
        for node in el.children:
            if node.name == 'br':
                p.addText('\n')
            else:
                process_text_node(node, p)

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

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(content)

    messages = []

    # DOCX conversion using python-docx with highlight preservation
    try:
        with open(html_path, 'r', encoding='utf-8') as fh:
            html = fh.read()
        try:
            html_to_docx(html, docx_path)
            messages.append('DOCX generado (preservando highlights)')
        except Exception as e:
            messages.append('No se pudo generar DOCX: ' + str(e))
    except Exception as e:
        messages.append('Error leyendo HTML para DOCX: ' + str(e))

    # ODT conversion using odfpy + BeautifulSoup with highlight styles
    try:
        with open(html_path, 'r', encoding='utf-8') as fh:
            html = fh.read()
        try:
            html_to_odt(html, odt_path)
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
                html_to_docx(html, target_path)

            elif ext == 'odt':
                with open(html_path, 'r', encoding='utf-8') as fh:
                    html = fh.read()
                html_to_odt(html, target_path)

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

    return send_file(target_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
