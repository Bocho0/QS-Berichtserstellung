"""
Gemeinsamer Kern der QS-Berichtserstellung - wird sowohl vom lokalen
CLI-Skript (build_report.py) als auch von der Vercel-Serverfunktion
(api/generate.py) verwendet, damit es nur EINE gepflegte Quelle für die
Formatierungslogik gibt.
"""
import os, copy, base64, re
import docx
from docx.shared import Pt, Emu
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree
from xml.sax.saxutils import escape as xml_escape
from PIL import Image

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
EMU_PER_IN = 914400
MAX_BOX_IN = 2.5
LINE_HEIGHT_PT = 9 * 1.2


def w(tag):
    return f'{{{W}}}{tag}'


def decode_photo(data_url, out_path):
    header, b64 = data_url.split(',', 1)
    with open(out_path, 'wb') as f:
        f.write(base64.b64decode(b64))
    return out_path


def fit_box(w_px, h_px, max_in=MAX_BOX_IN):
    ratio = w_px / h_px
    if ratio >= 1:
        w_in = max_in
        h_in = max_in / ratio
    else:
        h_in = max_in
        w_in = max_in * ratio
    return Emu(int(w_in * EMU_PER_IN)), Emu(int(h_in * EMU_PER_IN))


def set_run_font(run, base_run):
    run.font.name = base_run.font.name or 'Barlow'
    run.font.size = Pt(9)
    if base_run.font.color and base_run.font.color.rgb:
        run.font.color.rgb = base_run.font.color.rgb


def clear_cell(cell):
    tc = cell._tc
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)
    if len(cell.paragraphs) == 0:
        tc.append(parse_xml(f'<w:p xmlns:w="{W}"/>'))


def add_line(cell, text, base_run):
    p = cell.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, base_run)
    return p


def format_datum(raw):
    raw = (raw or '').strip()
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', raw)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}.{int(mo):02d}.{y}"
    return raw


WEEKDAYS_DE = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']


def format_datum_mit_wochentag(raw):
    """z.B. '08.09.2026' -> 'Di.08.09.2026'"""
    raw = (raw or '').strip()
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', raw)
    if not m:
        return raw
    d, mo, y = (int(x) for x in m.groups())
    try:
        import datetime
        wd = WEEKDAYS_DE[datetime.date(y, mo, d).weekday()]
        return f"{wd}.{d:02d}.{mo:02d}.{y}"
    except Exception:
        return f"{d:02d}.{mo:02d}.{y}"


def build_text_cell(cell, entry, nr, base_run, image_height_pt):
    clear_cell(cell)
    first = cell.paragraphs[0]
    first.text = ''
    first_run = first.add_run(f"Gewerk: {entry.get('gewerk','')}")
    set_run_font(first_run, base_run)
    first.paragraph_format.space_before = Pt(16)
    line_count = 1

    add_line(cell, f"Ort: {entry.get('ort','')}", base_run); line_count += 1
    if entry.get('bauteil'):
        add_line(cell, f"Bauteil: {entry['bauteil']}", base_run); line_count += 1
    add_line(cell, '', base_run); line_count += 1
    add_line(cell, 'Info:', base_run); line_count += 1
    info = entry.get('arbeiten', '') or ''
    for line in info.split('\n'):
        add_line(cell, line, base_run); line_count += 1
    if entry.get('material'):
        add_line(cell, f"Material: {entry['material']}", base_run); line_count += 1
    if entry.get('type') == 'veranlassung' and entry.get('verantwortlich'):
        add_line(cell, f"Verantwortlichkeit: {entry['verantwortlich']}", base_run); line_count += 1

    FINE_TUNE_PT = 6
    used_pt = 16 + line_count * LINE_HEIGHT_PT
    space_before = max(14, image_height_pt + 16 - used_pt - LINE_HEIGHT_PT - FINE_TUNE_PT)
    p_stand = add_line(cell, f"Stand:{format_datum(entry.get('datum',''))}", base_run)
    p_stand.paragraph_format.space_before = Pt(space_before)


def build_image_cell(cell, photo_path):
    clear_cell(cell)
    tcPr = cell._tc.get_or_add_tcPr()
    old_valign = tcPr.find(qn('w:vAlign'))
    if old_valign is not None:
        tcPr.remove(old_valign)
    valign = tcPr.makeelement(qn('w:vAlign'), {})
    valign.set(qn('w:val'), 'center')
    tcPr.append(valign)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(16)
    r = p.add_run()
    with Image.open(photo_path) as im:
        w_px, h_px = im.size
    w_emu, h_emu = fit_box(w_px, h_px)
    r.add_picture(photo_path, width=w_emu, height=h_emu)
    return h_emu / EMU_PER_IN * 72


def set_nr_cell(cell, nr, base_run):
    clear_cell(cell)
    tcPr = cell._tc.get_or_add_tcPr()
    old_valign = tcPr.find(qn('w:vAlign'))
    if old_valign is not None:
        tcPr.remove(old_valign)
    valign = tcPr.makeelement(qn('w:vAlign'), {})
    valign.set(qn('w:val'), 'top')
    tcPr.append(valign)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(16)
    # Die Kopfzeile "Nr.:" nutzt die Formatvorlage "Listenabsatz", die selbst
    # einen fest eingebauten Einzug von 720 dxa (0,5") mitbringt (in der
    # Absatz-XML selbst nicht sichtbar, da nur geerbt). Die Datenzeilen nutzen
    # dagegen "Normal" mit einem eigenen, davon abweichenden Einzug (360 dxa)
    # - das erzeugte den Versatz. Hier wird der Einzug der Datenzeile exakt
    # auf denselben Wert (720 dxa = 36pt) gesetzt, den die Kopfzeile durch
    # ihre Formatvorlage bereits hat.
    p.paragraph_format.left_indent = Pt(36)
    r = p.add_run(str(nr))
    set_run_font(r, base_run)


def replace_value_after_label(doc, label_start, new_value):
    for p in doc.paragraphs:
        if p.text.startswith(label_start):
            non_ul_runs = [r for r in p.runs if not r.font.underline]
            if not non_ul_runs:
                continue
            full_after = ''.join(r.text for r in non_ul_runs)
            m = re.match(r'^[\t ]*', full_after)
            prefix = m.group(0) if m else ''
            non_ul_runs[0].text = prefix + new_value
            for r in non_ul_runs[1:]:
                r.text = ''
            return True
    return False


def set_seitenanzahl_field(doc):
    """Setzt statt einer festen Zahl ein echtes Word-Feld (NUMPAGES) ein,
    damit die Seitenanzahl auch nach spaeteren Bearbeitungen in Word
    automatisch stimmt - kein Renderer/LibreOffice zur Vorab-Berechnung
    noetig (wichtig fuer die serverlose Umgebung ohne LibreOffice)."""
    for p in doc.paragraphs:
        if p.text.startswith('Seitenanzahl:'):
            non_ul_runs = [r for r in p.runs if not r.font.underline]
            if not non_ul_runs:
                return
            full_after = ''.join(r.text for r in non_ul_runs)
            m = re.match(r'^[\t ]*', full_after)
            prefix = m.group(0) if m else ''
            # Formatierung von einem bereits korrekt formatierten Datentext im
            # selben Absatz uebernehmen (z. B. " Seiten"), damit die eingefuegte
            # Zahl exakt dieselbe Schriftart/-groesse wie der uebrige Text nutzt,
            # statt eine eigene Schriftart fest zu verdrahten.
            ref_run = non_ul_runs[-1]
            ref_rPr = ref_run._r.find(qn('w:rPr'))
            if ref_rPr is not None:
                rpr_xml = etree.tostring(ref_rPr, encoding='unicode')
            else:
                rpr_xml = f'<w:rPr xmlns:w="{W}"><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/></w:rPr>'
            non_ul_runs[0].text = prefix
            for r in non_ul_runs[1:]:
                r.text = ''
            def make_run(inner):
                return parse_xml(f'<w:r xmlns:w="{W}">{rpr_xml}{inner}</w:r>')
            anchor = non_ul_runs[0]._r
            for rn in [
                make_run('<w:fldChar w:fldCharType="begin"/>'),
                make_run('<w:instrText>NUMPAGES   \\* MERGEFORMAT</w:instrText>'),
                make_run('<w:fldChar w:fldCharType="separate"/>'),
                make_run('<w:t>1</w:t>'),
                make_run('<w:fldChar w:fldCharType="end"/>'),
                make_run('<w:t xml:space="preserve"> Seiten</w:t>'),
            ]:
                anchor.addnext(rn)
                anchor = rn
            return


def set_col_width_dxa(table, col_idx, dxa):
    for row in table.rows:
        if col_idx < len(row.cells):
            tc = row.cells[col_idx]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = tcPr.makeelement(qn('w:tcW'), {})
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(dxa))
            tcW.set(qn('w:type'), 'dxa')
    grid = table._tbl.tblGrid
    cols = grid.findall(qn('w:gridCol'))
    if col_idx < len(cols):
        cols[col_idx].set(qn('w:w'), str(dxa))


def clear_data_rows(table, header_rows=1):
    for row in table.rows[header_rows:]:
        for cell in row.cells:
            if cell.text.strip() in ('Termin unkritisch', 'Termin kritisch', 'Termin überschritten'):
                continue
            for p in cell.paragraphs:
                for r in p.runs:
                    r.text = ''


def build(template_path, data, out_path, tmp_dir='/tmp/report_photos'):
    """data: bereits geparstes dict (nicht Dateipfad!)."""
    bk = data.get('berichtskopf', {})
    entries = data.get('entries', [])
    verlauf = data.get('verlauf', [])
    start_nr = data.get('startNr') or 1
    try:
        start_nr = int(start_nr)
    except (TypeError, ValueError):
        start_nr = 1

    doc = docx.Document(template_path)

    # Referenz auf die Dokumentationstabelle SOFORT einfangen, bevor weiter
    # unten neue Tabellen/Absätze (Inhaltsverzeichnis, Verlauf) in den
    # Dokumentkörper eingefügt werden - das würde sonst den Index in
    # doc.tables verschieben, sodass doc.tables[3] später auf die falsche
    # (neu eingefügte) Tabelle zeigen würde statt auf die Dokumentationstabelle.
    dokumentation_table = doc.tables[3]

    # Titelüberschrift auf eine Zeile bringen
    for p in doc.paragraphs:
        if p.text.strip().startswith('LEISTUNGSFESTELLUNG'):
            for r in p.runs:
                rPr = r._r.get_or_add_rPr()
                spacing_el = rPr.find(qn('w:spacing'))
                if spacing_el is not None:
                    spacing_el.set(qn('w:val'), '30')
                if r.font.size:
                    r.font.size = Pt(14)
            break

    # Inhaltsverzeichnis auf dem Titelblatt einfügen. Da die tatsächliche
    # Seitenaufteilung erst beim Öffnen in Word feststeht (abhängig von der
    # Anzahl/Länge der Feststellungen), wird hier - genau wie bei der
    # Seitenzahl - ein echtes Word-Feld (TOC) eingesetzt statt fester Werte.
    # Es befüllt sich automatisch beim Öffnen (updateFields ist weiter unten
    # bereits aktiviert); sollte Word das Feld nicht von selbst aktualisieren,
    # genügt ein Rechtsklick darauf → "Felder aktualisieren".
    def set_outline_level(paragraph, level):
        pPr = paragraph._p.get_or_add_pPr()
        old = pPr.find(qn('w:outlineLvl'))
        if old is not None:
            pPr.remove(old)
        el = pPr.makeelement(qn('w:outlineLvl'), {})
        el.set(qn('w:val'), str(level))
        pPr.append(el)

    anchor_p = None
    pagebreak_p = None
    heading_ref_rPr = None
    dokumentation_p = None
    leistung_p = None
    thema_p = None
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith('LEISTUNGSFESTELLUNG'):
            leistung_p = p
        if t.startswith('Thema:') and thema_p is None:
            thema_p = p
        if t.startswith('Anlagen'):
            anchor_p = p
        if t in ('Rahmentermine:', 'Dokumentation:'):
            set_outline_level(p, 0)
            if t == 'Dokumentation:' and p.runs:
                heading_ref_rPr = p.runs[0]._r.find(qn('w:rPr'))
                dokumentation_p = p
        if pagebreak_p is None:
            for br in p._p.findall('.//' + qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    pagebreak_p = p
                    break

    # Etwas Luft auf dem Titelblatt zurückgewinnen, damit für das weiter
    # unten eingefügte Inhaltsverzeichnis verlässlich Platz bleibt (auch bei
    # einer vollen Verteilerliste): den recht großzügigen Leerraum vor
    # "Thema:" sowie die ungenutzten Leerzeilen zwischen "Anlagen:" und dem
    # Seitenumbruch reduzieren. Betrifft nur leere Absätze, keine Inhalte.
    if leistung_p is not None and thema_p is not None:
        blanks = []
        p = leistung_p._p.getnext()
        while p is not None and p is not thema_p._p:
            blanks.append(p)
            p = p.getnext()
        for extra in blanks[3:]:
            extra.getparent().remove(extra)
    if anchor_p is not None and pagebreak_p is not None:
        p = anchor_p._p.getnext()
        while p is not None and p is not pagebreak_p._p:
            nxt = p.getnext()
            if not ''.join(p.itertext()).strip():
                p.getparent().remove(p)
            p = nxt

    if anchor_p is not None and pagebreak_p is not None:
        ref_rpr_xml = (etree.tostring(heading_ref_rPr, encoding='unicode')
                       if heading_ref_rPr is not None
                       else f'<w:rPr xmlns:w="{W}"><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/></w:rPr>')

        toc_heading = parse_xml(f'''<w:p xmlns:w="{W}">
          <w:pPr><w:spacing w:before="160" w:after="120"/></w:pPr>
          <w:r>{ref_rpr_xml}<w:t>Inhaltsverzeichnis</w:t></w:r>
        </w:p>''')

        toc_field = parse_xml(f'''<w:p xmlns:w="{W}">
          <w:pPr><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:noProof/></w:rPr>
            <w:fldChar w:fldCharType="begin" w:dirty="true"/>
          </w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            <w:instrText xml:space="preserve"> TOC \\o "1-1" \\h \\z \\u </w:instrText>
          </w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:noProof/></w:rPr>
            <w:fldChar w:fldCharType="separate"/>
          </w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:i/><w:noProof/></w:rPr>
            <w:t>Wird beim Öffnen automatisch befüllt (ggf. Rechtsklick → Felder aktualisieren).</w:t>
          </w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:noProof/></w:rPr>
            <w:fldChar w:fldCharType="end"/>
          </w:r>
        </w:p>''')

        pagebreak_p._p.addprevious(toc_heading)
        pagebreak_p._p.addprevious(toc_field)

    # Verlauf vorheriger Berichte (Kurzdarstellung) direkt über der
    # Dokumentationstabelle einfügen - als echte Nr.-Bereich/Zeitraum/
    # Zusammenfassung-Tabelle (analog zur Vorschau in der App), damit frühere
    # Feststellungen als kompakte Referenz sichtbar bleiben, ohne bei jedem
    # neuen Bericht der Reihe erneut als volle Einträge (samt Fotos)
    # aufzutauchen.
    if verlauf and dokumentation_p is not None:
        vref_rpr_xml = (etree.tostring(heading_ref_rPr, encoding='unicode')
                        if heading_ref_rPr is not None
                        else f'<w:rPr xmlns:w="{W}"><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/></w:rPr>')

        verlauf_heading = parse_xml(f'''<w:p xmlns:w="{W}">
          <w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
          <w:r>{vref_rpr_xml}<w:t>Verlauf vorheriger Berichte</w:t></w:r>
        </w:p>''')
        dokumentation_p._p.addprevious(verlauf_heading)

        sec = doc.sections[0]
        avail_dxa = sec.page_width.twips - sec.left_margin.twips - sec.right_margin.twips
        col_nr, col_zeit = 1300, 2600
        col_zus = avail_dxa - col_nr - col_zeit

        def vcell(text, width, header=False):
            sz = '18' if not header else '18'
            bold = '<w:b/>' if header else ''
            fill = ' <w:shd w:val="clear" w:color="auto" w:fill="EDEBE6"/>' if header else ''
            return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{fill}</w:tcPr>'
                     f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/>{bold}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
                     f'<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:p></w:tc>')

        rows_xml = '<w:tr>' + vcell('Nr.', col_nr, True) + vcell('Zeitraum', col_zeit, True) + vcell('Zusammenfassung', col_zus, True) + '</w:tr>'
        for v in verlauf:
            von, bis = v.get('von', ''), v.get('bis', '')
            nr_text = str(von) if von == bis else f'{von}–{bis}'
            zvon, zbis = v.get('zeitraumVon', '') or '', v.get('zeitraumBis', '') or ''
            zeit_text = zvon if zvon == zbis else f'{zvon} – {zbis}'
            berichtnr = str(v.get('berichtNr', '') or '').zfill(3) if str(v.get('berichtNr', '')).isdigit() else str(v.get('berichtNr', '') or '')
            zus_text = f"QS-Bericht {berichtnr}: " + (v.get('zusammenfassung', '') or '')
            rows_xml += '<w:tr>' + vcell(nr_text, col_nr) + vcell(zeit_text, col_zeit) + vcell(zus_text, col_zus) + '</w:tr>'

        verlauf_tbl_xml = f'''<w:tbl xmlns:w="{W}">
          <w:tblPr>
            <w:tblW w:w="{avail_dxa}" w:type="dxa"/>
            <w:tblLayout w:type="fixed"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>
              <w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid><w:gridCol w:w="{col_nr}"/><w:gridCol w:w="{col_zeit}"/><w:gridCol w:w="{col_zus}"/></w:tblGrid>
          {rows_xml}
        </w:tbl>'''
        verlauf_tbl = parse_xml(verlauf_tbl_xml)
        dokumentation_p._p.addprevious(verlauf_tbl)

        verlauf_spacer = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>')
        dokumentation_p._p.addprevious(verlauf_spacer)

    # Fußzeile: Dateiname (links) und Seitenzahl (rechts) auf eine gemeinsame
    # Zeile bringen. Beides lag bisher in zwei getrennten Absätzen unter-
    # einander - dafür wird hier eine randlose 2-Spalten-Tabelle aufgebaut,
    # die garantiert beide Angaben auf derselben Höhe zeigt. Das bestehende
    # Seitenzahl-Feld (PAGE/NUMPAGES) wird dabei unverändert weiterverwendet,
    # nur in die rechte Zelle verschoben.
    ftr = doc.sections[0].footer
    sdt = ftr._element.find(f'.//{w("sdt")}')
    dateiname_val = bk.get('dateiname') or (
        (bk.get('datum', '').replace('.', '') or 'bericht') + '_' +
        (bk.get('verfasser', '') or 'QS') + '-QS-Bautenstand_' +
        str(bk.get('berichtsNr', '1')).zfill(3) + '.docx')
    if sdt is not None:
        sdt_content = sdt.find(qn('w:sdtContent'))
        sdt_paragraphs = sdt_content.findall(qn('w:p')) if sdt_content is not None else []
        if sdt_paragraphs:
            target_p = sdt_paragraphs[-1]
            rpr_xml = f'<w:rPr xmlns:w="{W}"><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
            def make_run(inner):
                return parse_xml(f'<w:r xmlns:w="{W}">{rpr_xml}{inner}</w:r>')
            target_p.append(make_run('<w:t xml:space="preserve"> / </w:t>'))
            target_p.append(make_run('<w:fldChar w:fldCharType="begin"/>'))
            target_p.append(make_run('<w:instrText>NUMPAGES   \\* MERGEFORMAT</w:instrText>'))
            target_p.append(make_run('<w:fldChar w:fldCharType="separate"/>'))
            target_p.append(make_run('<w:t>1</w:t>'))
            target_p.append(make_run('<w:fldChar w:fldCharType="end"/>'))
            # Die leere erste Zeile innerhalb der Seitenzahl-Steuerelements
            # entfernen, damit die rechte Zelle nachher nur EINE Zeile hoch
            # ist (sonst wuerde die Tabellenzeile hoeher als die linke Zelle
            # und Dateiname/Seitenzahl liegen wieder auf unterschiedlicher
            # Hoehe).
            for extra_p in sdt_paragraphs[:-1]:
                sdt_content.remove(extra_p)

        # Alte, separate Fußzeilen-Absätze außerhalb des Seitenzahl-Steuer-
        # elements entfernen - der Dateiname zieht stattdessen gleich in die
        # neue Tabelle unten.
        for p_el in list(ftr._element.findall(qn('w:p'))):
            ftr._element.remove(p_el)

        sec = doc.sections[0]
        avail_dxa = sec.page_width.twips - sec.left_margin.twips - sec.right_margin.twips
        left_dxa = int(avail_dxa * 0.65)
        right_dxa = avail_dxa - left_dxa

        sdt.getparent().remove(sdt)

        tbl_xml = f'''<w:tbl xmlns:w="{W}">
          <w:tblPr>
            <w:tblW w:w="{avail_dxa}" w:type="dxa"/>
            <w:tblLayout w:type="fixed"/>
            <w:tblBorders>
              <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
              <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
              <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
              <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
              <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
              <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:w="0" w:type="dxa"/>
              <w:left w:w="0" w:type="dxa"/>
              <w:bottom w:w="0" w:type="dxa"/>
              <w:right w:w="0" w:type="dxa"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w="{left_dxa}"/>
            <w:gridCol w:w="{right_dxa}"/>
          </w:tblGrid>
          <w:tr>
            <w:tc>
              <w:tcPr><w:tcW w:w="{left_dxa}" w:type="dxa"/></w:tcPr>
              <w:p>
                <w:pPr><w:jc w:val="left"/></w:pPr>
                <w:r>
                  <w:rPr><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
                  <w:t xml:space="preserve">{dateiname_val}</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr><w:tcW w:w="{right_dxa}" w:type="dxa"/></w:tcPr>
            </w:tc>
          </w:tr>
        </w:tbl>'''
        tbl = parse_xml(tbl_xml)
        right_tc = tbl.findall(qn('w:tr') + '/' + qn('w:tc'))[1]
        right_tc.append(sdt)
        ftr._element.append(tbl)

    # Spaltenbreiten-Korrektur (Nr.-Spalte / Bautenstand-Spalte)
    try:
        set_col_width_dxa(dokumentation_table, 0, 1250)
        set_col_width_dxa(dokumentation_table, 1, 5068)
        set_col_width_dxa(dokumentation_table, 2, 3888)
    except Exception:
        pass
    try:
        set_col_width_dxa(doc.tables[2], 1, 2150)
        set_col_width_dxa(doc.tables[2], 4, 1780)
    except Exception:
        pass

    # Titelblatt
    replace_value_after_label(doc, 'Verfasser:', bk.get('verfasser', ''))
    replace_value_after_label(doc, 'Datum:', bk.get('datum', ''))
    replace_value_after_label(doc, 'Kunde:', bk.get('kunde', ''))
    set_seitenanzahl_field(doc)
    for p in doc.paragraphs:
        if p.text.strip() in ('Mustermannstr. 1', '12345 Hausen'):
            for r in p.runs:
                r.text = ''

    # Rahmentermine-/Wetter-Beispieldaten leeren. header_rows=0 bei der
    # Wettertabelle, da dort - anders als bei den anderen Tabellen - keine
    # echte Kopfzeile existiert; alle 3 Zeilen enthalten Datum/Temperatur
    # (nur die Legenden-Beschriftung je Zeile bleibt automatisch erhalten,
    # da clear_data_rows Zellen mit exaktem Label-Text ausspart).
    clear_data_rows(doc.tables[1], header_rows=0)
    clear_data_rows(doc.tables[2])

    # Witterung (Datum/Min./Max.) in die Wetter-/Farblegenden-Tabelle
    # eintragen - bisher wurden diese Werte zwar in der App gesammelt, aber
    # nie tatsächlich ins Word-Dokument übernommen.
    wetter_tbl = doc.tables[1]
    datum_mit_wochentag = format_datum_mit_wochentag(bk.get('datum', ''))
    temp_min = (bk.get('tempMin') or '').strip()
    temp_max = (bk.get('tempMax') or '').strip()
    for row in wetter_tbl.rows:
        if datum_mit_wochentag and len(row.cells) > 0:
            cell = row.cells[0]
            for p in cell.paragraphs:
                if p.runs:
                    p.runs[0].text = datum_mit_wochentag
                    for r in p.runs[1:]:
                        r.text = ''
        if temp_min and len(row.cells) > 1:
            cell = row.cells[1]
            value_p = cell.paragraphs[1] if len(cell.paragraphs) > 1 else None
            if value_p is not None and value_p.runs:
                value_p.runs[0].text = temp_min
                for r in value_p.runs[1:]:
                    r.text = ''
        if temp_max and len(row.cells) > 2:
            cell = row.cells[2]
            value_p = cell.paragraphs[1] if len(cell.paragraphs) > 1 else None
            if value_p is not None and value_p.runs:
                value_p.runs[0].text = temp_max
                for r in value_p.runs[1:]:
                    r.text = ''

    # Rahmentermine-Tabelle befüllen (bisher wurde dieses Feld gesammelt aber
    # nie tatsächlich in die Word-Tabelle geschrieben - Format je Zeile:
    # "Bezeichnung; Terminplan; Bautenstand %; Status"). Die App kennt vier
    # Status-Werte (im-termin/nicht-kritisch/plus2/plus4), die Word-Vorlage
    # aber nur drei Ampelfarben - im-termin und nicht-kritisch werden daher
    # beide auf Grün gemappt.
    if bk.get('rahmentermine', '').strip():
        RAHMEN_COLOR = {
            'im-termin': 'B2CB7F',
            'nicht-kritisch': 'B2CB7F',
            'plus2': 'F8A764',
            'plus4': 'F95649',
        }
        rt = doc.tables[2]
        rt_base_run = None
        for p in rt.rows[0].cells[1].paragraphs:
            if p.runs:
                rt_base_run = p.runs[0]
                break
        rows_text = [r.strip() for r in bk['rahmentermine'].split('\n') if r.strip()]
        for i, row_text in enumerate(rows_text):
            if i + 1 >= len(rt.rows):
                break
            parts = [p.strip() for p in row_text.split(';')]
            parts += [''] * (4 - len(parts))
            bezeichnung, terminplan, bautenstand, status = parts[:4]
            row = rt.rows[i + 1]
            col_text = {0: bezeichnung, 2: terminplan, 4: bautenstand}
            for col_idx, val in col_text.items():
                if not val or col_idx >= len(row.cells):
                    continue
                cell = row.cells[col_idx]
                for p in cell.paragraphs:
                    for r in list(p.runs):
                        r.text = ''
                target_p = cell.paragraphs[0]
                r = target_p.add_run(val)
                if rt_base_run is not None:
                    set_run_font(r, rt_base_run)
            fill = RAHMEN_COLOR.get(status.lower())
            if fill and len(row.cells) > 6:
                color_cell = row.cells[6]
                tcPr = color_cell._tc.get_or_add_tcPr()
                old_shd = tcPr.find(qn('w:shd'))
                if old_shd is not None:
                    tcPr.remove(old_shd)
                shd = tcPr.makeelement(qn('w:shd'), {})
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), fill)
                tcPr.append(shd)

    # Kopfzeile: Projektname (Logo wird nie angefasst). Nur das Projekt-Feld
    # wird angezeigt - eine Verkettung mit der "Baustellenbezeichnung" führte
    # zu einer sich wiederholenden Darstellung des Baustellen-Kürzels (z. B.
    # "NFM | München — Neufreimann WA11"), da beide Felder faktisch denselben
    # Standort beschreiben.
    projekt_text = bk.get('projekt', '').strip()
    for sec in doc.sections:
        for hdr in (sec.header, sec.first_page_header):
            for p in hdr.paragraphs:
                if 'XY' in p.text:
                    text_runs = [r for r in p.runs if r._r.find(qn('w:drawing')) is None]
                    drawing_runs = [r for r in p.runs if r._r.find(qn('w:drawing')) is not None]
                    for extra in drawing_runs[1:]:
                        extra._r.getparent().remove(extra._r)
                    if not text_runs:
                        continue
                    text_runs[0].text = projekt_text
                    for r in text_runs[1:]:
                        if r.text.strip() and 'Leistungsfeststellung' not in r.text and 'Qualitätssicherung' not in r.text:
                            r.text = ''

    # Verteiler-Tabelle
    if bk.get('verteiler', '').strip():
        vt = doc.tables[0]
        vt_base_run = vt.rows[0].cells[0].paragraphs[0].runs[0]
        header_tcs = vt.rows[0]._tr.findall(qn('w:tc'))
        rows_text = [r.strip() for r in bk['verteiler'].split('\n') if r.strip()]
        for i, row_text in enumerate(rows_text):
            if i + 1 >= len(vt.rows):
                break
            parts = [p.strip() for p in row_text.split(';')]
            parts += [''] * (4 - len(parts))
            values = parts[:4]
            row_tr = vt.rows[i + 1]._tr
            for old_tc in list(row_tr.findall(qn('w:tc'))):
                row_tr.remove(old_tc)
            for htc, val in zip(header_tcs, values):
                new_tc = copy.deepcopy(htc)
                for p_el in new_tc.findall(qn('w:p')):
                    new_tc.remove(p_el)
                new_tc.append(parse_xml(f'<w:p xmlns:w="{W}"/>'))
                tcPr = new_tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    for b in tcPr.findall(qn('w:tcBorders')):
                        tcPr.remove(b)
                    tcPr.append(parse_xml(
                        f'<w:tcBorders xmlns:w="{W}">'
                        '<w:top w:val="dotted" w:sz="2" w:space="0" w:color="auto"/>'
                        '<w:bottom w:val="dotted" w:sz="2" w:space="0" w:color="auto"/>'
                        '</w:tcBorders>'))
                row_tr.append(new_tc)
                cell = [c for c in vt.rows[i + 1].cells if c._tc is new_tc][0]
                r = cell.paragraphs[0].add_run(val)
                set_run_font(r, vt_base_run)
        # Ungenutzte, leere Vorlagenzeilen entfernen (die Tabelle ist auf 12
        # Datenzeilen ausgelegt, unabhängig davon wie viele tatsächlich
        # gebraucht werden) - das schafft verlässlich Freiraum auf dem
        # Titelblatt, u. a. damit das Inhaltsverzeichnis dort sicher Platz
        # hat, auch wenn Word es später auf mehrere Zeilen erweitert.
        used_rows = min(len(rows_text), len(vt.rows) - 1)
        for row in list(vt.rows[used_rows + 1:]):
            vt._tbl.remove(row._tr)

    # Dokumentationstabelle
    table = dokumentation_table
    tbl = table._tbl
    template_row_tr = copy.deepcopy(table.rows[1]._tr)
    base_run = table.rows[1].cells[1].paragraphs[1].runs[0]

    for r in list(table.rows)[1:]:
        tbl.remove(r._tr)

    os.makedirs(tmp_dir, exist_ok=True)
    for i, entry in enumerate(entries, start=start_nr):
        new_tr = copy.deepcopy(template_row_tr)
        tbl.append(new_tr)
        row = table.rows[-1]
        set_nr_cell(row.cells[0], i, base_run)
        photo_path = os.path.join(tmp_dir, f"entry_{i}.jpg")
        decode_photo(entry['photoUrl'], photo_path)
        image_height_pt = build_image_cell(row.cells[2], photo_path)
        build_text_cell(row.cells[1], entry, i, base_run, image_height_pt)

    doc.save(out_path)

    # Word anweisen, Felder (Seitenzahl-Felder etc.) beim Öffnen automatisch
    # neu zu berechnen, statt den eingebetteten Platzhalterwert stehen zu lassen.
    import zipfile
    tmp_fixed = out_path + '.tmp'
    with zipfile.ZipFile(out_path, 'r') as zin:
        names = zin.namelist()
        settings_xml = zin.read('word/settings.xml').decode('utf-8')
        if '<w:updateFields' not in settings_xml:
            settings_xml = settings_xml.replace(
                '<w:settings',
                '<w:settings', 1)
            settings_xml = settings_xml.replace(
                '</w:settings>',
                '<w:updateFields w:val="true"/></w:settings>', 1)
        with zipfile.ZipFile(tmp_fixed, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_bytes = settings_xml.encode('utf-8') if item.filename == 'word/settings.xml' else zin.read(item.filename)
                zout.writestr(item, data_bytes)
    os.replace(tmp_fixed, out_path)
    return out_path
