# -*- coding: utf-8 -*-
import os, sys, io, json, zipfile, tempfile, threading, webbrowser
from flask import Flask, request, send_file

def ruta_base():
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

os.chdir(ruta_base())
sys.path.insert(0, ruta_base())

import generar_resoluciones as gen

app = Flask(__name__)

@app.after_request
def sin_cache(resp):
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

def limpiar(nombre):
    nombre = (nombre or 'resolucion').strip()
    for ch in '\\/:*?"<>|':
        nombre = nombre.replace(ch, '_')
    return nombre or 'resolucion'

def procesar(registros):
    tmpdir = tempfile.mkdtemp(prefix='mce_')
    hechos = []
    for reg in registros:
        nombre = limpiar(reg.get('nombre_est'))
        salida = os.path.join(tmpdir, nombre + '.docx')
        n = 2
        while salida in [h[1] for h in hechos]:
            salida = os.path.join(tmpdir, nombre + '_' + str(n) + '.docx')
            n += 1
        if gen.generar(reg, salida) and os.path.exists(salida):
            hechos.append((os.path.basename(salida), salida))
    return hechos

@app.route('/')
def index():
    with open('formulario_resoluciones_MCE.html', 'r', encoding='utf-8') as f:
        return f.read()

def _responder():
    raw = request.get_data(as_text=True)
    if request.form.get('datos'):
        raw = request.form.get('datos')
    datos = json.loads(raw)
    registros = datos if isinstance(datos, list) else [datos]

    hechos = procesar(registros)
    if not hechos:
        return "Error: no se genero ningun documento. Verifique que exista la plantilla para ese tramite.", 500

    if len(hechos) == 1:
        nom, ruta = hechos[0]
        with open(ruta, 'rb') as f:
            data = f.read()
        return send_file(io.BytesIO(data),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True, download_name=nom)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for nom, ruta in hechos:
            z.write(ruta, nom)
    buf.seek(0)
    return send_file(buf, mimetype='application/zip',
        as_attachment=True, download_name='resoluciones.zip')

@app.route('/version')
def version():
    return {'version': 'v7-OK', 'archivo': os.path.basename(__file__),
            'rutas': sorted(str(r) for r in app.url_map.iter_rules())}

@app.route('/api/generar', methods=['POST'])
def api_generar():
    try:
        return _responder()
    except Exception as e:
        import traceback; traceback.print_exc()
        return "Error: " + str(e), 500

@app.route('/descargar', methods=['POST'])
def descargar():
    try:
        return _responder()
    except Exception as e:
        import traceback; traceback.print_exc()
        return "Error: " + str(e), 500

if __name__ == '__main__':
    def abrir():
        import time; time.sleep(1.5)
        webbrowser.open('http://127.0.0.1:5000')
    threading.Thread(target=abrir, daemon=True).start()
    print("Servidor en http://127.0.0.1:5000")
    app.run(debug=False, host='127.0.0.1', port=5000)
