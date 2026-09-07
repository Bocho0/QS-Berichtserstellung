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
from PIL import Image

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
EMU_PER_IN = 914400
MAX_BOX_IN = 3.6
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
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(16)
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
            non_ul_runs[0].text = prefix
            for r in non_ul_runs[1:]:
                r.text = ''
            rpr_xml = f'<w:rPr xmlns:w="{W}"><w:rFonts w:ascii="Barlow" w:hAnsi="Barlow"/></w:rPr>'
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

    doc = docx.Document(template_path)

    # Titelüberschrift auf eine Zeile bringen
    for p in doc.paragraphs:
        if p.text.strip().startswith('LEISTUNGSFESTELLUNG'):
            for r in p.runs:
                rPr = r._r.get_or_add_rPr()
                spacing_el = rPr.find(qn('w:spacing'))
                if spacing_el is not None:
                    spacing_el.set(qn('w:val'), '30')
                if r.font.size:
                    r.font.size = Pt(12)
            break

    # Fußzeile: Gesamtseitenzahl + Dateiname ergänzen (Seitenzahl-Feld war
    # schon vorhanden)
    ftr = doc.sections[0].footer
    sdt = ftr._element.find(f'.//{w("sdt")}')
    if sdt is not None:
        sdt_paragraphs = sdt.findall(f'.//{w("p")}')
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
        trailing_paragraphs = [p for p in ftr.paragraphs if p._p.getparent() is ftr._element]
        dateiname_val = bk.get('dateiname') or (
            (bk.get('datum', '').replace('.', '') or 'bericht') + '_' +
            (bk.get('verfasser', '') or 'QS') + '-QS-Bautenstand_' +
            str(bk.get('berichtsNr', '1')).zfill(3) + '.docx')
        if trailing_paragraphs:
            tp = trailing_paragraphs[-1]
            tp.text = ''
            r = tp.add_run(dateiname_val)
            r.font.name = 'Barlow'
            r.font.size = Pt(9)

    # Spaltenbreiten-Korrektur (Nr.-Spalte / Bautenstand-Spalte)
    try:
        set_col_width_dxa(doc.tables[3], 0, 1250)
        set_col_width_dxa(doc.tables[3], 1, 3712)
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

    # Rahmentermine-/Wetter-Beispieldaten leeren
    clear_data_rows(doc.tables[1])
    clear_data_rows(doc.tables[2])

    # Kopfzeilen: Projektname (Logo wird nie angefasst)
    projekt_text = (bk.get('projekt', '') + (' — ' + bk['baustelle'] if bk.get('baustelle') else '')).strip(' —')
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

    # Dokumentationstabelle
    table = doc.tables[3]
    tbl = table._tbl
    template_row_tr = copy.deepcopy(table.rows[1]._tr)
    base_run = table.rows[1].cells[1].paragraphs[1].runs[0]

    for r in list(table.rows)[1:]:
        tbl.remove(r._tr)

    os.makedirs(tmp_dir, exist_ok=True)
    for i, entry in enumerate(entries, start=1):
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
