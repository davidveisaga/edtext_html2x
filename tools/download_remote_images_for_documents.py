#!/usr/bin/env python3
"""
Download remote <img src="http(s):..."> referenced in HTML files inside `documents/`.
Saves images to `static/uploads/` and rewrites the HTML to reference the local copy
using `../static/uploads/<file>` so they display when opening the HTML and so
conversions embed them.
"""
import os
import re
import urllib.request
import mimetypes
import imghdr
import uuid

DOCS = 'documents'
UPLOAD = os.path.join('static', 'uploads')
os.makedirs(UPLOAD, exist_ok=True)

updated = 0
for fname in os.listdir(DOCS):
    if not fname.endswith('.html'):
        continue
    path = os.path.join(DOCS, fname)
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    changed = False
    # find all img src="http..."
    matches = re.findall(r'<img[^>]+src=["\'](http[^"\']+)["\']', html)
    for url in matches:
        try:
            with urllib.request.urlopen(url) as resp:
                data = resp.read()
                info = resp.info()
                ctype = info.get_content_type() if hasattr(info, 'get_content_type') else None
            ext = None
            if ctype:
                ext = mimetypes.guess_extension(ctype)
            if not ext:
                t = imghdr.what(None, h=data)
                if t:
                    ext = '.' + t
            if not ext:
                ext = '.bin'
            fname_local = f"{uuid.uuid4().hex}{ext}"
            fpath = os.path.join(UPLOAD, fname_local)
            with open(fpath, 'wb') as fh:
                fh.write(data)
            # replace only first occurrence for this URL
            html = html.replace(url, f"../static/uploads/{fname_local}", 1)
            print(f"{path}: downloaded {url} -> {fpath}")
            changed = True
        except Exception as e:
            print(f"Failed to download {url}: {e}")
    if changed:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
        updated += 1

print('Done. Files updated:', updated)
