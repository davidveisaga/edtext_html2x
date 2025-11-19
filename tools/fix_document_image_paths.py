#!/usr/bin/env python3
"""
Normalize image src paths in HTML files under `documents/` to use
`../static/uploads/...` so they work both when served by Flask and when
opened directly from the filesystem.
"""
import os
import re

DOCS = 'documents'
count = 0
for fname in os.listdir(DOCS):
    if not fname.endswith('.html'):
        continue
    path = os.path.join(DOCS, fname)
    with open(path, 'r', encoding='utf-8') as f:
        s = f.read()
    orig = s
    # Replace /static/uploads/ -> ../static/uploads/
    s = re.sub(r'src=["\']?/static/uploads/', 'src="../static/uploads/', s)
    # Replace file://.../static/uploads/ -> ../static/uploads/
    s = re.sub(r'src=["\']?file://[^"\']*?/static/uploads/', 'src="../static/uploads/', s)
    if s != orig:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(s)
        print('Updated', path)
        count += 1

print('Done. Files updated:', count)
