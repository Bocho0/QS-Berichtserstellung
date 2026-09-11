"""
Vercel Python-Funktion /api/weather - fragt Geokodierung + historische
Temperaturwerte bei Open-Meteo ab und reicht sie durch.

Grund fuer diesen Umweg ueber den eigenen Server: Direkte Browser-Aufrufe zu
geocoding-api.open-meteo.com scheiterten mit einem CORS-Fehler ("No
'Access-Control-Allow-Origin' header"). Server-zu-Server-Aufrufe sind davon
nicht betroffen (CORS ist ausschliesslich eine Browser-Einschraenkung), daher
uebernimmt diese Funktion die Anfrage stellvertretend fuer die App.
"""
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
import json, traceback
import urllib.request
import urllib.error


def _fetch_json(url, timeout=6):
    req = urllib.request.Request(url, headers={'User-Agent': 'qs-bautenstand-app/1.0'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode('utf-8'))


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def do_GET(self):
        try:
            qs = parse_qs(urlparse(self.path).query)
            ort = (qs.get('ort', [''])[0] or '').strip()
            datum = (qs.get('datum', [''])[0] or '').strip()  # Format YYYY-MM-DD

            if not ort or not datum:
                self._send_json(400, {'error': 'Parameter "ort" und "datum" (JJJJ-MM-TT) werden benötigt.'})
                return

            geo_url = ('https://geocoding-api.open-meteo.com/v1/search?name='
                       + quote(ort) + '&count=1&language=de&format=json')
            try:
                geo = _fetch_json(geo_url)
            except (urllib.error.URLError, TimeoutError) as e:
                self._send_json(502, {'error': f'Geokodierung fehlgeschlagen: {e}'})
                return

            results = geo.get('results') or []
            if not results:
                self._send_json(404, {'error': f'Ort "{ort}" nicht gefunden.'})
                return

            lat = results[0]['latitude']
            lon = results[0]['longitude']
            name = results[0].get('name', ort)

            weather_url = (
                'https://archive-api.open-meteo.com/v1/archive'
                f'?latitude={lat}&longitude={lon}&start_date={datum}&end_date={datum}'
                '&daily=temperature_2m_max,temperature_2m_min&timezone=auto'
            )
            try:
                wdata = _fetch_json(weather_url)
            except (urllib.error.URLError, TimeoutError) as e:
                self._send_json(502, {'error': f'Wetterabruf fehlgeschlagen: {e}'})
                return

            daily = wdata.get('daily') or {}
            tmax_list = daily.get('temperature_2m_max') or []
            tmin_list = daily.get('temperature_2m_min') or []
            tmax = tmax_list[0] if tmax_list else None
            tmin = tmin_list[0] if tmin_list else None

            if tmax is None or tmin is None:
                self._send_json(404, {'error': 'Für dieses Datum liegen keine Wetterdaten vor.'})
                return

            self._send_json(200, {'ort': name, 'tempMin': round(tmin, 1), 'tempMax': round(tmax, 1)})
        except Exception as e:
            traceback.print_exc()
            self._send_json(500, {'error': str(e)})

    def _cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _send_json(self, code, obj):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self._cors_headers()
        self.end_headers()
        self.wfile.write(json.dumps(obj, ensure_ascii=False).encode('utf-8'))
