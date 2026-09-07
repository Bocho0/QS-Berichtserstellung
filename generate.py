"""
Vercel Python-Funktion: erzeugt aus den von der qs-bautenstand.html App
gesendeten Daten den fertigen Word-Bericht und liefert ihn direkt als
Download zurueck. Nutzt dieselbe Formatierungslogik wie das lokale
CLI-Skript (report_core.py).
"""
import os, json, traceback
from flask import Flask, request, send_file, jsonify
import report_core

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template', 'vorlage.docx')

@app.route('/api/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return _cors(app.make_response(''))
    try:
        data = request.get_json(force=True)
        if not data or not data.get('entries'):
            return _cors(jsonify({'error': 'Keine Einträge in den gesendeten Daten gefunden.'})), 400

        out_path = '/tmp/bericht.docx'
        report_core.build(TEMPLATE_PATH, data, out_path, tmp_dir='/tmp/report_photos')

        bk = data.get('berichtskopf', {})
        dateiname = bk.get('dateiname') or (
            (bk.get('datum', '').replace('.', '') or 'bericht') + '_' +
            (bk.get('verfasser', '') or 'QS') + '-QS-Bautenstand_' +
            str(bk.get('berichtsNr', '1')).zfill(3) + '.docx')

        resp = send_file(
            out_path,
            as_attachment=True,
            download_name=dateiname,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        return _cors(resp)
    except Exception as e:
        traceback.print_exc()
        return _cors(jsonify({'error': str(e)})), 500

def _cors(resp):
    resp.headers['Access-Control-Allow-Origin'] = '*'
    resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp
