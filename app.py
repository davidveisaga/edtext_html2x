from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
import subprocess
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'documents'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
    grouped = {}

    for f in files:
        if f.endswith(('.html', '.odt', '.docx')):
            base = f.rsplit('.', 1)[0]
            ext = f.rsplit('.', 1)[1]
            if base not in grouped:
                grouped[base] = []
            grouped[base].append(ext)

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
    docname = request.form['docname'].strip().replace(' ', '_')
    filename = secure_filename(request.form.get('filename') or f"{docname}.html")

    html_path = os.path.join(UPLOAD_FOLDER, filename)
    odt_path = html_path.replace('.html', '.odt')
    docx_path = html_path.replace('.html', '.docx')

    with open(html_path, 'w') as f:
        f.write(content)

    try:
        subprocess.run(["pandoc", html_path, "-o", odt_path], check=True)
        subprocess.run(["pandoc", html_path, "-o", docx_path], check=True)
        flash("✅ Documento guardado con éxito.")
    except subprocess.CalledProcessError:
        flash("❌ Error al convertir el documento con Pandoc.")

    return render_template('form.html', content=content, filename=filename)

@app.route('/documents/<filename>')
def download(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
