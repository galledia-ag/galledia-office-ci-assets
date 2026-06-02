"""
fill_dokument.py — Galledia Dokument-Skill v1.1
CI-konforme .docx aus Vorlage_Dokument.dotx (V3)

build_document(
    titel, untertitel, datum, rechtseinheit, adresse, empfaenger,
    abschnitte=[{
        'titel': 'Kapitel',
        'inhalt': [
            {'typ': 'text',    'inhalt': 'Fliesstext'},
            {'typ': 'bullet',  'inhalt': 'Aufzählung'},
            {'typ': 'h2',      'inhalt': 'Unterkapitel'},
            {'typ': 'h3',      'inhalt': 'Abschnitt'},
            {'typ': 'tabelle', 'inhalt': [['Kopf A','Kopf B'],['Wert','Wert']]},
        ]
    }],
    output_path='output.docx'
)
"""
import io, os, zipfile
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_TEMPLATE = os.path.join(os.path.dirname(__file__), "assets", "Vorlage_Dokument.dotx")

# ── Template-Loader ───────────────────────────────────────────────────────────

def _load(path=None):
    """Lädt .dotx als editierbares Document (patcht Content_Types.xml im ZIP)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(path or _TEMPLATE, 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    data = data.replace(
                        b'.wordprocessingml.template.main+xml',
                        b'.wordprocessingml.document.main+xml')
                zout.writestr(item, data)
    buf.seek(0)
    return Document(buf)

# ── Primitive ─────────────────────────────────────────────────────────────────

def _set_text(para, text):
    """Ersetzt Paragraph-Text; \\n-Zeichen → Word-Zeilenumbrüche."""
    for r in para.runs:
        r.text = ''
    lines = str(text).split('\n')
    if para.runs:
        para.runs[0].text = lines[0]
    else:
        para.add_run(lines[0])
    last = para.runs[0]._r
    for line in lines[1:]:
        br = OxmlElement('w:br')
        last.addnext(br); last = br
        r  = OxmlElement('w:r')
        t  = OxmlElement('w:t'); t.text = line; r.append(t)
        last.addnext(r); last = r

def _para(style, text=''):
    """Erstellt ein neues Paragraphen-Element mit Style und Text."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    ps  = OxmlElement('w:pStyle'); ps.set(qn('w:val'), style)
    pPr.append(ps); p.append(pPr)
    if text:
        r = OxmlElement('w:r')
        t = OxmlElement('w:t'); t.text = text
        if text != text.strip():
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t); p.append(r)
    return p

# ── TOC ───────────────────────────────────────────────────────────────────────

def _toc(ref_p):
    """
    Fügt TOC auf eigener Seite ein:
    pageBreakBefore auf der Überschrift → TOC auf neuer Seite.
    Seitenumbruch nach TOC-Feld → Inhalt auf neuer Seite.
    """
    def _fc(ftype):
        r = OxmlElement('w:r'); fc = OxmlElement('w:fldChar')
        fc.set(qn('w:fldCharType'), ftype); r.append(fc); return r

    # TOC-Überschrift mit pageBreakBefore
    h = OxmlElement('w:p')
    hPr = OxmlElement('w:pPr')
    hStyle = OxmlElement('w:pStyle')
    hStyle.set(qn('w:val'), 'Inhaltsverzeichnisberschrift')
    hPr.append(hStyle)
    h.append(hPr)   # Seitenumbruch via Template-Sektionsstruktur
    hR = OxmlElement('w:r'); hT = OxmlElement('w:t')
    hT.text = 'Inhaltsverzeichnis'; hR.append(hT); h.append(hR)
    ref_p._p.addnext(h)

    # TOC-Feld (Word aktualisiert mit F9)
    tp = OxmlElement('w:p')
    tpPr = OxmlElement('w:pPr'); tpStyle = OxmlElement('w:pStyle')
    tpStyle.set(qn('w:val'), 'Verzeichnis1'); tpPr.append(tpStyle); tp.append(tpPr)
    tp.append(_fc('begin'))
    ri = OxmlElement('w:r'); ins = OxmlElement('w:instrText')
    ins.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    ins.text = ' TOC \\o "1-3" \\h \\z \\u '; ri.append(ins); tp.append(ri)
    tp.append(_fc('separate'))
    rt = OxmlElement('w:r'); tt = OxmlElement('w:t')
    tt.text = 'Inhaltsverzeichnis — in Word mit F9 aktualisieren.'
    rt.append(tt); tp.append(rt)
    tp.append(_fc('end'))
    h.addnext(tp)

    # Seitenumbruch nach TOC → Inhalt auf neuer Seite
    pb = OxmlElement('w:p')
    pbR = OxmlElement('w:r'); pbBr = OxmlElement('w:br')
    pbBr.set(qn('w:type'), 'page'); pbR.append(pbBr); pb.append(pbR)
    tp.addnext(pb)

    return pb   # Anchor: Content-Insertion startet nach diesem Element

# ── Tabelle ───────────────────────────────────────────────────────────────────

def _table(rows_data):
    """
    CI-konforme Tabelle.
    Kopfzeile: Volte Semibold, Schwarz, Hintergrund #F2F2F5.
    Datenzeilen: Volte Regular.
    """
    tbl = OxmlElement('w:tbl')
    tblPr = OxmlElement('w:tblPr')
    tW = OxmlElement('w:tblW')
    tW.set(qn('w:w'), '5000'); tW.set(qn('w:type'), 'pct')
    tblPr.append(tW); tbl.append(tblPr)

    for ri, row in enumerate(rows_data):
        tr = OxmlElement('w:tr')
        for cell in row:
            tc  = OxmlElement('w:tc')
            tcPr = OxmlElement('w:tcPr')
            if ri == 0:
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'),   'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'),  'F2F2F5')
                tcPr.append(shd)
            tc.append(tcPr)

            cp  = OxmlElement('w:p')
            cr  = OxmlElement('w:r')
            crPr = OxmlElement('w:rPr')

            # Font: Volte Semibold (Kopf) / Volte Regular (Daten)
            fn = OxmlElement('w:rFonts')
            fn.set(qn('w:ascii'), 'Volte Semibold' if ri == 0 else 'Volte')
            fn.set(qn('w:hAnsi'), 'Volte Semibold' if ri == 0 else 'Volte')
            crPr.append(fn)

            # Kopfzeile: Schwarz, kein Fett-Tag (Semibold reicht)
            col = OxmlElement('w:color')
            col.set(qn('w:val'), '000000')  # immer Schwarz
            crPr.append(col)

            sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '18'); crPr.append(sz)

            cr.append(crPr)
            ct = OxmlElement('w:t'); ct.text = str(cell)
            cr.append(ct); cp.append(cr); tc.append(cp); tr.append(tc)
        tbl.append(tr)
    return tbl

# ── Hauptfunktion ─────────────────────────────────────────────────────────────

def build_document(
    titel,
    untertitel    = '',
    datum         = '',
    rechtseinheit = '',
    adresse       = '',    # \n für Zeilenumbrüche: 'Strasse\nPLZ Ort'
    empfaenger    = '',    # optional, \n für Zeilenumbrüche
    abschnitte    = None,
    output_path   = 'output.docx',
    template      = None
):
    """
    Erstellt ein CI-konformes Galledia-Dokument aus Vorlage_Dokument.dotx.

    Aufbau:  Deckblatt | Inhaltsverzeichnis (eigene Seite) | Inhalt
    Schluss: Ort/Datum und Rechtseinheit stehen auf dem Deckblatt — nicht wiederholt.

    abschnitte = [{'titel': str, 'inhalt': [{'typ': str, 'inhalt': str|list}]}]
    typ-Werte:  'text' | 'bullet' | 'h2' | 'h3' | 'tabelle'
    """
    doc = _load(template)

    # ── 1. Leere Inhaltsverzeichnis-Überschriften entfernen (Vorlage-Artefakte) ─
    for p in list(doc.paragraphs):
        if 'Inhaltsverzeichnis' in p.style.name and not p.text.strip():
            p._p.getparent().remove(p._p)

    # ── 2. Deckblatt-Platzhalter befüllen ────────────────────────────────────
    datum_set = False
    for p in doc.paragraphs:
        t = p.text.strip()
        if   t == '[Dokumenttitel]':
            _set_text(p, titel)
        elif t == '[Untertitel / Anlass]':
            _set_text(p, untertitel)
        elif '[Ort], [Datum]' in t and not datum_set:
            _set_text(p, datum); datum_set = True
        elif '[Rechtseinheit] / [Adresse]' in t:
            _set_text(p, rechtseinheit + ('\n' + adresse if adresse else ''))
        elif t == '[Empfänger-Block, optional]':
            _set_text(p, empfaenger)

    # ── 3. Fusszeile: [Dokumenttitel] → Titel ────────────────────────────────
    for sect in doc.sections:
        for hf in (sect.footer, getattr(sect, 'first_page_footer', None)):
            if hf is None:
                continue
            for fp in hf.paragraphs:
                if '[Dokumenttitel]' in fp.text:
                    for run in fp.runs:
                        run.text = run.text.replace('[Dokumenttitel]', titel)

    # ── 4. TOC einfügen (eigene Seite) ───────────────────────────────────────
    toc_end = None
    for p in doc.paragraphs:
        if p.text.strip() == '[Inhaltsverzeichnis]':
            toc_end = _toc(p)
            p._p.getparent().remove(p._p)
            break

    # ── 5. Inhalt einfügen ───────────────────────────────────────────────────
    anchor = None
    for p in doc.paragraphs:
        if p.text.strip() == '[Text / Inhalt]':
            anchor = p; break

    if anchor and abschnitte:
        last = anchor._p
        for si, section in enumerate(abschnitte):
            if si > 0:   # Leerzeile zwischen Kapiteln
                sp = _para('Standard', ''); last.addnext(sp); last = sp
            h1 = _para('berschrift1', section.get('titel', ''))
            last.addnext(h1); last = h1

            for item in section.get('inhalt', []):
                typ = item.get('typ', 'text')
                cnt = item.get('inhalt', '')
                style_map = {
                    'text':   'Standard',
                    'bullet': 'DokBullet',
                    'h2':     'berschrift2',
                    'h3':     'berschrift3',
                }
                if typ in ('h2', 'h3'):   # Leerzeile vor Unterüberschriften
                    sp = _para('Standard', ''); last.addnext(sp); last = sp
                if typ == 'tabelle':
                    tbl = _table(cnt)
                    last.addnext(tbl); last = tbl
                    sp  = _para('Standard', ''); last.addnext(sp); last = sp
                else:
                    np = _para(style_map.get(typ, 'Standard'), cnt)
                    last.addnext(np); last = np

        anchor._p.getparent().remove(anchor._p)

    doc.save(output_path)
    return output_path
