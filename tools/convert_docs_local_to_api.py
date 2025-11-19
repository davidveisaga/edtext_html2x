#!/usr/bin/env python3
"""
Rewrite local upload references in HTML files under `documents/` to use
`/api/images/<filename>` so the app-served pages and converters use the API URLs.
"""
import os
import re

DOCS = 'documents'
changed = 0
for fname in os.listdir(DOCS):
    if not fname.endswith('.html'):
        continue
    path = os.path.join(DOCS, fname)
    with open(path, 'r', encoding='utf-8') as f:
        s = f.read()
    orig = s
    # Replace ../static/uploads/<file> or /static/uploads/<file> or file://.../static/uploads/<file>
    s = re.sub(r'src=["\']?\.\./static/uploads/([^"\'>]+)', lambda m: f'src="/api/images/{m.group(1)}"', s)
    s = re.sub(r'src=["\']?/static/uploads/([^"\'>]+)', lambda m: f'src="/api/images/{m.group(1)}"', s)
    s = re.sub(r'src=["\']?file://[^"\']*?/static/uploads/([^"\'>]+)', lambda m: f'src="/api/images/{m.group(1)}"', s)
    if s != orig:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(s)
        print('Updated', path)
        changed += 1
print('Done. Files updated:', changed)
