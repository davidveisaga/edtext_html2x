#!/usr/bin/env python3
"""
Script de prueba para la conversión HTML a DOCX.
Convierte un archivo HTML de prueba a DOCX y luego inspecciona el XML 
interno del archivo resultante para verificar la correcta implementación
de los highlights usando los elementos w:shd.
"""

import os
from app import html_to_docx
from zipfile import ZipFile

html_path = os.path.join('documents','01.html')
out_path = os.path.join('documents','01_test.docx')

with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

print('Generating', out_path)
html_to_docx(html, out_path)
print('Saved. Inspecting docx xml for shading elements...')

with ZipFile(out_path, 'r') as z:
    names = z.namelist()
    for n in names:
        if n.endswith('.xml') and (n.endswith('document.xml') or 'document' in n):
            print('---', n, '---')
            data = z.read(n).decode('utf-8')
            # print a window around shd occurrences
            idx = data.find('w:shd')
            if idx >= 0:
                start = max(0, idx-200)
                end = min(len(data), idx+200)
                print(data[start:end])
            else:
                print('No w:shd found in', n)

print('Done')
