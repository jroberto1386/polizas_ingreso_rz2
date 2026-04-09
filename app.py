"""
Plataforma de Pólizas de Ingreso — RZ2 Sistemas / GBC Business Consulting
Servidor Flask local — corre en http://localhost:5050
v2.0: soporte para archivo REP (Complementos de Pago)
"""

from flask import Flask, render_template, request, send_file, jsonify
import os, uuid, threading, time
from werkzeug.utils import secure_filename
from motor import procesar_polizas

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'outputs')
CATALOGO_BASE = os.path.join(os.path.dirname(__file__), 'Cat_Clientes_RZ2.xlsx')

jobs = {}


def limpiar_archivos_viejos():
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        for f in os.listdir(folder):
            fp = os.path.join(folder, f)
            if os.path.isfile(fp) and (time.time() - os.path.getmtime(fp)) > 7200:
                try:
                    os.remove(fp)
                except Exception:
                    pass


def run_job(job_id, ruta_banco, ruta_facturas, ruta_catalogo, ruta_reps):
    try:
        jobs[job_id].update(status='running', progress=20,
                            message='Leyendo extracto bancario...')

        output_path = os.path.join(OUTPUT_FOLDER, f'polizas_{job_id[:8]}.xlsx')

        jobs[job_id].update(progress=40, message='Leyendo facturas pendientes...')
        time.sleep(0.2)

        msg = 'Matching banco ↔ facturas ↔ REP...' if ruta_reps else 'Ejecutando matching banco ↔ facturas...'
        jobs[job_id].update(progress=55, message=msg)

        stats = procesar_polizas(
            ruta_banco, ruta_facturas, ruta_catalogo, output_path,
            ruta_reps=ruta_reps
        )

        jobs[job_id].update(progress=90, message='Convirtiendo a .xls para CONTPAq...')
        time.sleep(0.3)

        ruta_final = stats.get('ruta_final', output_path)
        es_xls = ruta_final.endswith('.xls') and not ruta_final.endswith('.xlsx')

        jobs[job_id].update(
            status='done', progress=100,
            message='¡Proceso completado!',
            file_path=ruta_final,
            stats={**stats, 'n_ppd_sin_rep': stats.get('n_sin_rep', 0)},
            es_xls=es_xls,
        )

    except Exception as e:
        jobs[job_id].update(status='error', message=f'Error: {str(e)}', progress=0)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    limpiar_archivos_viejos()
    job_id = str(uuid.uuid4())

    def guardar(key):
        f = request.files.get(key)
        if not f or f.filename == '':
            return None
        fname = f'{job_id[:8]}_{secure_filename(f.filename)}'
        path  = os.path.join(UPLOAD_FOLDER, fname)
        f.save(path)
        return path

    ruta_banco    = guardar('banco')
    ruta_facturas = guardar('facturas')
    ruta_catalogo = guardar('catalogo') or CATALOGO_BASE
    ruta_reps     = guardar('rep')    # opcional

    if not ruta_banco or not ruta_facturas:
        return jsonify({'error': 'Debes subir el extracto bancario y las facturas.'}), 400

    if not os.path.exists(ruta_catalogo):
        return jsonify({'error': 'No se encontró el catálogo de clientes. Sube el archivo Cat_Clientes_RZ2.xlsx.'}), 400

    jobs[job_id] = {
        'status': 'queued', 'progress': 5,
        'message': 'En cola...', 'file_path': None, 'stats': None,
        'con_reps': ruta_reps is not None,
    }

    t = threading.Thread(
        target=run_job,
        args=(job_id, ruta_banco, ruta_facturas, ruta_catalogo, ruta_reps)
    )
    t.daemon = True
    t.start()

    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job no encontrado'}), 404
    return jsonify(job)


@app.route('/descargar/<job_id>')
def descargar(job_id):
    job = jobs.get(job_id)
    if not job or not job.get('file_path'):
        return 'Archivo no disponible', 404
    ruta = job['file_path']
    es_xls = ruta.endswith('.xls') and not ruta.endswith('.xlsx')
    nombre = 'polizas_ingreso_CONTPAq.xls' if es_xls else 'polizas_ingreso_CONTPAq.xlsx'
    mime   = 'application/vnd.ms-excel' if es_xls else \
             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return send_file(ruta, as_attachment=True, download_name=nombre, mimetype=mime)


if __name__ == '__main__':
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    print("\n  ✅  Plataforma de Pólizas de Ingreso v2.0")
    print("  🌐  Abre tu navegador en: http://localhost:5050\n")
    app.run(host='0.0.0.0', port=5050, debug=False)
