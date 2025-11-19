"""Normalize documents in `documents/`:
- add `data-local-src` for any `api/images/<file>` or `static/uploads/<file>` img
- ensure `src` uses absolute API URL (`/api/images/<file>`) when possible
- inject file:// fallback script if `data-local-src` present

Usage: python tools/normalize_documents.py
"""
import os
import shutil
from bs4 import BeautifulSoup
import re
DOCS = os.path.join(os.path.dirname(__file__), '..', 'documents')
DOCS = os.path.normpath(DOCS)
BACKUP = os.path.join(DOCS, 'backup')
os.makedirs(BACKUP, exist_ok=True)

def rewrite_content(original):
    """Rewrite img srcs to use absolute API path and add data-local-src."""
    try:
        soup = BeautifulSoup(original, 'lxml')
    except Exception:
        soup = BeautifulSoup(original, 'html.parser')
    changed = False
    for img in soup.find_all('img'):
        src = img.get('src', '')
        if not src:
            continue
        # static/uploads/filename
        m = re.search(r'static/uploads/([^"\'>]+)', src)
        if m:
            filename = m.group(1)
            img['src'] = f"/api/images/{filename}"
            img['data-local-src'] = f"../static/uploads/{filename}"
            changed = True
            continue
        # api/images/filename (may be ../api/images/...)
        m2 = re.search(r'api/images/([^"\'>\\]+)', src)
        if m2:
            filename = m2.group(1)
            img['src'] = f"/api/images/{filename}"
            img['data-local-src'] = f"../static/uploads/{filename}"
            changed = True
            continue

    new = str(soup)
    # inject fallback script if needed
    if 'data-local-src' in new and '<!-- FILE_VIEW_FALLBACK_SCRIPT -->' not in new:
        marker = '<!-- FILE_VIEW_FALLBACK_SCRIPT -->'
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
        import re as _re
        if _re.search(r'</body>', new, flags=_re.IGNORECASE):
            new = _re.sub(r'</body>', script + '</body>', new, flags=_re.IGNORECASE)
        else:
            new = new + script
        changed = True

    return new, changed

changed_files = []
for fn in os.listdir(DOCS):
    if not fn.endswith('.html'):
        continue
    path = os.path.join(DOCS, fn)
    if os.path.isdir(path):
        continue
    with open(path, 'r', encoding='utf-8') as f:
        original = f.read()
    # backup
    shutil.copy2(path, os.path.join(BACKUP, fn))
    new, changed = rewrite_content(original)
    if changed and new != original:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new)
        changed_files.append(fn)

print(f"Normalized {len(changed_files)} files: {changed_files}")
