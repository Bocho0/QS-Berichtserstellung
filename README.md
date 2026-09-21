# QS-Berichtserstellung

Web-App zur Baustellen-QS-Berichtserstellung: Fotoaufnahme vor Ort, Erfassung der
Feststellungen/Veranlassungen, Generierung eines formalen Word-Berichts (.docx) im
Corporate-Design-Layout ("QS-Bericht Muster").

## Struktur

- `public/index.html` – Frontend (eine Datei, inline HTML/CSS/JS, IndexedDB-Persistenz)
- `api/index.py` – Vercel Python-Serverless-Funktion (`POST /api/generate` bzw. `/api`)
- `api/report_core.py` – befüllt die Vorlage mit den Berichtsdaten (Werte, Fotos,
  Tabellenzeilen, Inhaltsverzeichnis, Fußzeile). **Enthält bewusst keine
  Formatierungslogik mehr** - Schriften/Farben/Ränder/Linien kommen komplett aus
  `vorlage.docx`.
- `api/vorlage.docx` – Word-Vorlage. Enthält bereits die komplette Muster-Formatierung
  fest eingebaut (Schriften, Farben, Linien statt Rahmen, Spaltenbreiten, verbundene
  Kopfzellen, Wetterband, Ampel-Legende). Wird von `report_core.py` nur noch mit
  Werten befüllt, nicht mehr umformatiert.
- `scripts/build_template.py` – das Skript, mit dem `vorlage.docx` aus der
  ursprünglichen Rohvorlage erzeugt wurde. Nur relevant, falls die Formatierung
  grundlegend geändert werden soll (siehe unten).

## Architekturentscheidung

Frühere Version dieses Projekts erzeugte die komplette Muster-Formatierung bei
JEDER Berichtserstellung zur Laufzeit per Code. Das war fehleranfällig und schwer
nachvollziehbar. Jetzt gilt:

- **`vorlage.docx` bestimmt das Aussehen** (Schriften, Farben, Ränder, Linien,
  Spaltenbreiten). Wer die Formatierung ändern will, kann das größtenteils direkt
  in Word an `vorlage.docx` tun.
- **`report_core.py` befüllt nur noch Werte** (Text, Fotos, Tabellenzeilen klonen).
  Ausnahmen sind zwangsläufig dynamische Teile, die von der tatsächlichen
  Seitenaufteilung abhängen: Inhaltsverzeichnis (PAGEREF-Felder), Verlauf
  vorheriger Berichte, Fußzeilen-Seitenzahl (NUMPAGES-Feld).
- Tabellen werden über den Text ihrer Kopfzeile gefunden (`find_table_by_header`),
  nicht über eine feste Position in `doc.tables` - unempfindlich gegenüber
  zusätzlichen Tabellen in der Vorlage (z. B. der Ampel-Legende).
- Für die Dokumentations- und Rahmentermine-Tabelle enthält `vorlage.docx` genau
  EINE fertig formatierte Musterzeile, die pro Eintrag geklont wird (wie schon
  vorher bei der Dokumentationstabelle üblich).
- Für die Feststellungszeilen (Gewerk/Ort/Bauteil/Info/…) enthält die Vorlage
  einen Prototyp-Absatz ("PROTO_LABEL"/"PROTO_VALUE") mit dem fertigen
  Label/Wert-Raster (Tabulator bei 15mm, Hängeeinzug, Farben) - der Code
  dupliziert diesen Absatz je benötigtem Feld und ersetzt nur den Text.

## `scripts/build_template.py` erneut ausführen

Nötig nur, wenn die GRUNDFORMATIERUNG geändert werden soll (z. B. andere Farben,
andere Spaltenbreiten). Für normale Textkorrekturen (z. B. "Verteilung: per
E-Mail" statt "per Email") reicht es, `vorlage.docx` direkt in Word zu öffnen und
zu speichern.

```bash
pip install python-docx lxml
python3 scripts/build_template.py <ausgangs-vorlage.docx> api/vorlage.docx
```

## Deployment (Vercel)

1. Repository auf GitHub pushen.
2. In Vercel „Import Project" → GitHub-Repo auswählen.
3. Vercel erkennt `vercel.json` automatisch (Python-Funktion mit `api/vorlage.docx`
   als `includeFiles`). Kein Build-Schritt nötig, `public/` wird als statisches
   Verzeichnis ausgeliefert.
4. Word-Generierung läuft über `POST /api/generate`.

## Lokal testen

```bash
pip install -r requirements.txt
python3 -c "
import sys; sys.path.insert(0, 'api')
import report_core, json
data = json.load(open('testdata.json'))
report_core.build('api/vorlage.docx', data, 'output.docx', tmp_dir='/tmp/report_photos')
"
```
