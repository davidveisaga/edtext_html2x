#!/usr/bin/env python3
"""
Herramienta de inspección para archivos ODT.
Permite examinar el contenido XML dentro de archivos ODT, específicamente:
- Muestra el contenido de content.xml y styles.xml
- Lista los estilos automáticos definidos
- Encuentra y lista los spans con sus estilos
- Útil para debuggear la implementación de highlights en ODT
"""

import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

# Helper to strip namespace
def qname(tag):
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def dump_odt(odt_path):
    print(f"Inspecting: {odt_path}\n")
    with ZipFile(odt_path, 'r') as z:
        names = z.namelist()
        for n in ('content.xml', 'styles.xml'):
            if n in names:
                print(f"--- {n} (first 2000 chars) ---")
                data = z.read(n).decode('utf-8')
                print(data[:2000])
                print('\n')
        # Parse content.xml to find automatic styles and spans
        if 'content.xml' in names:
            content = z.read('content.xml').decode('utf-8')
            root = ET.fromstring(content)
            # collect namespace map
            ns = {}
            for k, v in root.attrib.items():
                if k.startswith('xmlns:'):
                    ns[k.split(':',1)[1]] = v
            print('Namespaces detected:', ns)
            # find automatic-styles
            auto = root.find('.//{http://openoffice.org/2000/office}automatic-styles')
            if auto is None:
                print('No automatic-styles in content.xml')
            else:
                print('\nAutomatic styles in content.xml:')
                for s in auto:
                    name = s.get('{http://www.w3.org/XML/1998/namespace}id') or s.get('name') or s.get('style:name') or s.get('style:name')
                    # try multiple ways
                    # fallback to attribute 'name'
                    if not name:
                        name = s.get('name') or s.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
                    if not name:
                        name = s.get('style:name') or s.get('style:name')
                    # get text-properties child
                    tp = None
                    for ch in s:
                        if qname(ch.tag).endswith('text-properties'):
                            tp = ch
                            break
                    if name is None:
                        # try to get style:name from attributes
                        name = s.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') or s.attrib.get('name')
                    print(' - style element tag=', qname(s.tag), 'attrs=', s.attrib)
                    if tp is not None:
                        print('   text-properties attrs=', tp.attrib)
                print('\n')
            # find all spans that have text:style-name
            spans = root.findall('.//')
            used = set()
            for el in root.iter():
                if qname(el.tag) == 'span' or qname(el.tag) == 'span':
                    # try attributes for text:style-name
                    for k, v in el.attrib.items():
                        if k.endswith('style-name') or k.endswith('styleName'):
                            used.add(v)
            print('Referenced span style-names (sample):')
            for u in sorted(list(used))[:200]:
                print(' -', u)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python inspect_odt.py <path/to/file.odt>')
        sys.exit(1)
    dump_odt(sys.argv[1])
