"""
Vercel Python-Funktion (einfaches Modell: eine Datei = eine Route
/api/generate). Erzeugt aus den von der App gesendeten Daten den fertigen
Word-Bericht und liefert ihn direkt als Download zurueck.
"""
from http.server import BaseHTTPRequestHandler
import json, os, traceback
import report_core

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template', 'vorlage.docx')
if not os.path.exists(TEMPLATE_PATH):
    for alt in [
        os.path.join(os.path.dirname(__file__), '..', 'api', 'template', 'vorlage.docx'),
        '/var/task/api/template/vorlage.docx',
        '/var/task/template/vorlage.docx',
    ]:
        if os.path.exists(alt):
            TEMPLATE_PATH = alt
            break


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self._cors_headers()
        self.end_headers()
        self.wfile.write('Diese Adresse nimmt nur POST-Anfragen zur Berichtserstellung entgegen.'.encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body)

            if not data or not data.get('entries'):
                self._send_json(400, {'error': 'Keine Einträge in den gesendeten Daten gefunden.'})
                return

            if not os.path.exists(TEMPLATE_PATH):
                self._send_json(500, {
                    'error': f'Vorlage nicht gefunden. Gesucht bei: {TEMPLATE_PATH}. '
                             f'Verzeichnisinhalt von {os.path.dirname(__file__)}: '
                             f'{os.listdir(os.path.dirname(__file__))}'
                })
                return

            out_path = '/tmp/bericht.docx'
            report_core.build(TEMPLATE_PATH, data, out_path, tmp_dir='/tmp/report_photos')

            bk = data.get('berichtskopf', {})
            dateiname = bk.get('dateiname') or (
                (bk.get('datum', '').replace('.', '') or 'bericht') + '_' +
                (bk.get('verfasser', '') or 'QS') + '-QS-Bautenstand_' +
                str(bk.get('berichtsNr', '1')).zfill(3) + '.docx'
            )

            with open(out_path, 'rb') as f:
                content = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{dateiname}"')
            self._cors_headers()
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            traceback.print_exc()
            self._send_json(500, {'error': str(e)})

    def _cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _send_json(self, code, obj):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self._cors_headers()
        self.end_headers()
        self.wfile.write(json.dumps(obj).encode('utf-8'))
