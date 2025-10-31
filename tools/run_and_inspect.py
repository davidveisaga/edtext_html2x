#!/usr/bin/env python3
"""
Script de prueba para la conversión HTML a ODT.
Convierte un archivo HTML de prueba a ODT y luego utiliza inspect_odt.py
para examinar el archivo resultante y verificar la correcta implementación
de los highlights y estilos.
"""

import os
from app import html_to_odt

BASE = os.path.join('documents', '12.html')
OUT = os.path.join('documents', '12_test.odt')

with open(BASE, 'r', encoding='utf-8') as f:
    html = f.read()

print('Running html_to_odt to generate', OUT)
html_to_odt(html, OUT)
print('Done conversion. Now inspect generated file...')

# reuse inspect_odt
import inspect_odt
inspect_odt.dump_odt(OUT)
