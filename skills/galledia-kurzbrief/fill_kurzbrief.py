"""
fill_kurzbrief.py — Galledia Kurzbrief (Code Execution, ohne MCP)

CI-konforme .docx aus Vorlage_Kurzbrief.dotx via Content-Controls (SDT).

Im Gegensatz zum Brief KEIN Body, sondern 10 vordefinierte Notiz-Optionen mit
Checkboxen, die im gedruckten Dokument manuell angekreuzt werden.

build_kurzbrief(
    sender_oe, sender_first_name, sender_last_name,
    sender_street, sender_zip, sender_city,
    sender_contact_email,
    sender_contact_phone=None, sender_contact_mobile=None,
    recipient_salutation, recipient_first_name, recipient_last_name,
    recipient_company=None, recipient_street, recipient_zip, recipient_city,
    date_city, date, subject,
    introduction=None, closing="Freundliche Grüsse",
    signatory_name=None, signatory_role=None,
    notes=None,  # dict[str, str] zum Überschreiben einzelner Standard-Notes
    output_path="kurzbrief.docx"
)
"""
import io, os, zipfile
from datetime import date as _date
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_TEMPLATE = os.path.join(os.path.dirname(__file__), "assets", "Vorlage_Kurzbrief.dotx")

_MONTHS_DE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
              "Juli", "August", "September", "Oktober", "November", "Dezember"]

def _today_de():
    """Aktuelles Datum im Galledia-Format '1. Juni 2026'."""
    t = _date.today()
    return f"{t.day}. {_MONTHS_DE[t.month - 1]} {t.year}"

# Standard-Notiz-Optionen (Markenhandbuch v1.5, Kurzbrief-Konvention)
DEFAULT_NOTES = {
    "Note1":  "zur Kenntnisnahme",
    "Note2":  "zu Ihren Akten",
    "Note3":  "auf Ihren Wunsch",
    "Note4":  "mit Dank zurück",
    "Note5":  "zur Erledigung",
    "Note6":  "gemäss telefonischer Besprechung",
    "Note7":  "zur Stellungnahme",
    "Note8":  "gemäss Ihrer Anfrage",
    "Note9":  "per E-Mail an:",
    "Note10": "Beilagen:",
}


def _load(path=None):
    """Lädt .dotx als editierbares .docx (patcht Content_Types.xml im ZIP)."""
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


def _find_sdts(doc):
    """Findet alle Content Controls (SDT) im Dokument."""
    sdts = {}
    body = doc.element.body
    for sdt in body.iter(qn('w:sdt')):
        sdtPr = sdt.find(qn('w:sdtPr'))
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(qn('w:tag'))
        if tag_el is None:
            continue
        tag_val = tag_el.get(qn('w:val'))
        if tag_val and tag_val not in sdts:
            sdts[tag_val] = sdt
    return sdts


from copy import deepcopy


def _make_run(text, rPr_template=None):
    """Erzeugt einen <w:r> mit Text und optional übernommenen Run-Properties."""
    r = OxmlElement('w:r')
    if rPr_template is not None:
        r.append(deepcopy(rPr_template))
    t = OxmlElement('w:t')
    t.text = text or ""
    if text and (text != text.strip() or text.startswith(" ")):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r


def _make_break_run(rPr_template=None):
    """Erzeugt einen <w:r><w:br/></w:r> mit optional übernommenen rPr."""
    r = OxmlElement('w:r')
    if rPr_template is not None:
        r.append(deepcopy(rPr_template))
    r.append(OxmlElement('w:br'))
    return r


def _first_rPr(parent):
    """Findet das erste <w:rPr> in einem <w:r>-Kind von parent (für Style-Übernahme)."""
    for r in parent.findall(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is not None:
            return rPr
    return None


def _add_text_runs(paragraph, text, multiline=False, rPr_template=None):
    """Fügt Text als Runs (mit \\n als <w:br/>) zu einem Paragraph hinzu.
    rPr_template wird in jeden Run kopiert (Font/Style übernehmen)."""
    if not text:
        paragraph.append(_make_run("", rPr_template))
        return
    lines = text.split('\n') if multiline else [text]
    for i, line in enumerate(lines):
        if i > 0:
            paragraph.append(_make_break_run(rPr_template))
        paragraph.append(_make_run(line, rPr_template))


def _clear_showing_placeholder(sdt):
    """Entfernt <w:showingPlcHdr/> aus <w:sdtPr> — sonst zeigt Word weiter den
    Glossary-Platzhalter statt unseres Contents."""
    sdtPr = sdt.find(qn('w:sdtPr'))
    if sdtPr is None:
        return
    plchdr = sdtPr.find(qn('w:showingPlcHdr'))
    if plchdr is not None:
        sdtPr.remove(plchdr)


def _set_sdt_text(sdt, text, multiline=False):
    """Setzt den Inhalt eines SDT auf den gegebenen Text.

    - Entfernt <w:showingPlcHdr/> (sonst rendert Word den Glossary-Platzhalter)
    - Übernimmt rPr (Font/Style) vom ersten existierenden Run
    - Inline-SDT (keine <w:p>): newlines werden als <w:br/> in einer Run-Sequenz
      direkt unter <w:sdtContent> ausgegeben (newlines in <w:t> würden zu Spaces
      kollabiert)
    """
    _clear_showing_placeholder(sdt)
    content = sdt.find(qn('w:sdtContent'))
    if content is None:
        return

    paragraphs = content.findall(qn('w:p'))
    if paragraphs:
        # Paragraph-SDT: rPr aus erstem Run des ersten Paragraphen übernehmen
        first_p = paragraphs[0]
        rPr_template = _first_rPr(first_p)
        for r in first_p.findall(qn('w:r')):
            first_p.remove(r)
        _add_text_runs(first_p, text, multiline=multiline, rPr_template=rPr_template)
        for p in paragraphs[1:]:
            content.remove(p)
    else:
        # Inline-SDT: rPr aus erstem direkten Run übernehmen
        rPr_template = _first_rPr(content)
        for r in content.findall(qn('w:r')):
            content.remove(r)
        if not text:
            content.append(_make_run("", rPr_template))
            return
        lines = text.split('\n') if multiline else [text]
        for i, line in enumerate(lines):
            if i > 0:
                content.append(_make_break_run(rPr_template))
            content.append(_make_run(line, rPr_template))


_OAW_NS = "http://schemas.officeatwork.com/CustomXMLPart"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _set_checkboxes_by_index(doc, indices_checked):
    """Setzt die ankreuzbaren Notes-Checkboxen (1..10) auf checked.

    Die Vorlage enthält genau 10 Checkbox-SDTs in der Reihenfolge Note1..Note10
    (ohne Tag-Namen). Wir iterieren in Dokument-Reihenfolge und kreuzen die
    übergebenen Indizes (1-basiert) an.

    indices_checked: iterable[int] — z.B. {1, 5, 10}
    """
    checked_set = set(indices_checked)
    if not checked_set:
        return
    body = doc.element.body
    idx = 0
    for sdt in body.iter(qn('w:sdt')):
        sdtPr = sdt.find(qn('w:sdtPr'))
        if sdtPr is None:
            continue
        cb = sdtPr.find(f'{{{_W14_NS}}}checkbox')
        if cb is None:
            continue
        idx += 1
        if idx not in checked_set:
            continue
        # 1) w14:checked auf 1 setzen
        checked_el = cb.find(f'{{{_W14_NS}}}checked')
        if checked_el is not None:
            checked_el.set(f'{{{_W14_NS}}}val', "1")
        # 2) sichtbaren Glyph in sdtContent von ☐ (U+2610) auf ☒ (U+2612)
        content = sdt.find(qn('w:sdtContent'))
        if content is not None:
            for t in content.iter(qn('w:t')):
                t.text = "☒"


def _patch_office_at_work(docx_path, values):
    """Patcht den officeatwork CustomXMLPart (customXml/itemN.xml) im
    gespeicherten docx, sodass Word die Werte beim Öffnen via <w:dataBinding>
    in alle SDTs schreibt. Ohne diesen Patch zeigt Word den Glossary-Default.

    values: dict[tag, text] — Tag-Namen wie Date, City, Note1, RecipientAddress.
    """
    from lxml import etree
    target_member = None
    new_xml_bytes = None

    with zipfile.ZipFile(docx_path, 'r') as zin:
        for name in zin.namelist():
            if not (name.startswith('customXml/item') and name.endswith('.xml')):
                continue
            if 'Props' in name or 'rels' in name:
                continue
            data = zin.read(name)
            if _OAW_NS.encode() not in data:
                continue
            target_member = name
            root = etree.fromstring(data)
            for child in list(root):
                local = etree.QName(child.tag).localname
                if local in values:
                    val = values[local]
                    if val is None:
                        continue
                    child.text = val
            new_xml_bytes = etree.tostring(
                root, xml_declaration=True, encoding='utf-8', standalone=False)
            break

    if target_member is None:
        # Kein officeatwork-Template — kein Patch nötig
        return

    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml_bytes if item.filename == target_member \
                    else zin.read(item.filename)
                zout.writestr(item, data)
    os.replace(tmp_path, docx_path)


def _full_address_line(salutation, first_name, last_name, company, street, zip_code, city):
    lines = []
    salut_str = f"{salutation} " if salutation else ""
    lines.append(f"{salut_str}{first_name} {last_name}".strip())
    if company:
        lines.append(company)
    lines.append(street)
    lines.append(f"{zip_code} {city}".strip())
    return "\n".join(lines)


def _sender_org_block(oe, street, zip_code, city):
    return f"{oe} · {street} · {zip_code} {city}"


def _doc_contact_block(first_name, last_name, email, phone=None, mobile=None):
    lines = [f"{first_name} {last_name}", email]
    if phone:
        lines.append(phone if phone.startswith("T ") else f"T {phone}")
    if mobile:
        lines.append(mobile if mobile.startswith("M ") else f"M {mobile}")
    return "\n".join(lines)


def _signature_block(name, role=None):
    if role:
        return f"{name}\n{role}"
    return name


def build_kurzbrief(
    sender_oe,
    sender_first_name, sender_last_name,
    sender_street, sender_zip, sender_city,
    sender_contact_email,
    recipient_salutation, recipient_first_name, recipient_last_name,
    recipient_street, recipient_zip, recipient_city,
    subject,
    date_city=None, date=None,
    sender_contact_phone=None, sender_contact_mobile=None,
    recipient_company=None,
    introduction=None,
    closing="Freundliche Grüsse",
    signatory_name=None, signatory_role=None,
    signatory_name_2=None, signatory_role_2=None,
    notes=None,
    checked=None,
    template=None,
    output_path="kurzbrief.docx",
):
    """Erstellt einen Galledia-Kurzbrief aus Vorlage_Kurzbrief.dotx.

    Args:
        sender_*, recipient_*, date_*, subject: wie bei build_brief
            (siehe fill_brief.py für Detail-Doku)
        notes: dict[str, str] zum Überschreiben einzelner Note-Defaults.
            Keys: "Note1" bis "Note10". JEDER hier vorhandene Key wird
            zusätzlich AUTOMATISCH ANGEKREUZT (☒).
            Beispiel: notes={"Note10": "Beilagen: Vertrag, AGB, NDA"}
                      → Note10 wird text-überschrieben UND angekreuzt.
        checked: optional list[str] — Notes, die NUR angekreuzt werden
            sollen (ohne Text-Override). Default-Text bleibt.
            Beispiel: checked=["Note1", "Note5"]

    Returns:
        output_path (str)
    """
    doc = _load(template)
    sdts = _find_sdts(doc)

    # Auto-Defaults
    if date is None:
        date = _today_de()
    if date_city is None:
        date_city = sender_city  # Ort = Absender-Standort

    # Default-Anrede
    if introduction is None:
        if recipient_last_name and recipient_salutation:
            sal = recipient_salutation.strip()
            if sal.lower().startswith("herr"):
                introduction = f"Sehr geehrter {sal} {recipient_last_name}"
            elif sal.lower().startswith("frau"):
                introduction = f"Sehr geehrte {sal} {recipient_last_name}"
            else:
                introduction = f"Sehr geehrte/r {sal} {recipient_last_name}"
        else:
            introduction = "Sehr geehrte Damen und Herren"

    # Default Signatory = Sender (name + role aus Args)
    if signatory_name is None:
        signatory_name = f"{sender_first_name} {sender_last_name}"

    # Notes-Dict aus Defaults + Overrides
    final_notes = dict(DEFAULT_NOTES)
    if notes:
        final_notes.update(notes)

    # AdressBlock-Komposition entspricht dem officeatwork-Original:
    # - Organisation       = nur Firmen-Name           ("Galledia Fachmedien AG")
    # - Address1           = Strasse                   ("Baslerstrasse 60")
    # - Address2           = PLZ + Ort                 ("8048 Zürich")
    # - AdressBlockOrganisation = nur Firmen-Name (gleich wie Organisation)
    # - AdressBlock        = Strasse + PLZ Ort, mehrzeilig
    # - Doc.Contact        = LABEL "Ihr Kontakt" (KEIN Block!)
    # - AdressBlockSignature1 = voller Kontaktblock (Name + T + M + Email)
    # - Signature1         = Name (+ optional Rolle)
    contact_block = _doc_contact_block(
        sender_first_name, sender_last_name, sender_contact_email,
        sender_contact_phone, sender_contact_mobile)
    sig1 = _signature_block(signatory_name, signatory_role)

    # Alle Felder dynamisch aus Args. AdressBlock (Sender, oberhalb Empfänger)
    # einzeilig: "Strasse · PLZ Ort" — multi-line würde mit dem Empfänger-Block
    # überlappen.
    mappings = {
        "Organisation": sender_oe,
        "Address1": sender_street,
        "Address2": f"{sender_zip} {sender_city}",
        # Multi-line wie Template-Default — mit Trailing-Newline für Leerzeile
        # vor "Ihr Kontakt" (Doc.Contact)
        "AdressBlock": f"{sender_street}\n{sender_zip} {sender_city}\n ",
        "AdressBlockOrganisation": sender_oe,
        "Doc.Contact": "Ihr Kontakt",
        "AdressBlockSignature1": contact_block,
        "RecipientAddress": _full_address_line(
            recipient_salutation, recipient_first_name, recipient_last_name,
            recipient_company, recipient_street, recipient_zip, recipient_city),
        "City": date_city,
        "Date": date,
        "Introduction": introduction,
        "Closing": closing,
        "Signature1": sig1,
        "Subject": subject,
        # Alle 10 Notes
        **final_notes,
    }
    if signatory_name_2:
        sig2 = _signature_block(signatory_name_2, signatory_role_2)
        mappings["Signature2"] = sig2
        mappings["AdressBlockSignature2"] = sig2

    multiline_keys = {
        "AdressBlock", "AdressBlockSignature1", "AdressBlockSignature2",
        "RecipientAddress", "Signature1", "Signature2",
    }
    # Belt-and-suspenders: setze SDT-Content UND patche Custom XML Part
    # (Word liest aus dem CustomXMLPart via <w:dataBinding>, aber falls
    # ein Renderer den CustomXMLPart ignoriert, ist der SDT-Content da)
    for tag, val in mappings.items():
        if tag in sdts:
            _set_sdt_text(sdts[tag], val, multiline=(tag in multiline_keys))

    # Checkboxes ankreuzen (alle Keys in notes-Dict + alle Keys in checked-Liste)
    indices_to_check = set()
    for src in (notes or {}, checked or []):
        for key in src:
            if isinstance(key, str) and key.startswith("Note"):
                try:
                    indices_to_check.add(int(key[4:]))
                except ValueError:
                    pass
    _set_checkboxes_by_index(doc, indices_to_check)

    # Subject als Titel (Fallback für Templates ohne Subject-SDT/-binding)
    if subject:
        for p in doc.element.body.iter(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None and pStyle.get(qn('w:val')) == "Titel":
                    for r in p.findall(qn('w:r')):
                        p.remove(r)
                    _add_text_runs(p, subject)
                    break

    doc.save(output_path)
    # Final pass: patch officeatwork CustomXMLPart (Bug-Fix v2.0.1)
    _patch_office_at_work(output_path, mappings)
    return output_path
