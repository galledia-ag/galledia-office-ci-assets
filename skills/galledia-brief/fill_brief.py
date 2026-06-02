"""
fill_brief.py — Galledia Geschäftsbrief (Code Execution, ohne MCP)

CI-konforme .docx aus Vorlage_Brief.dotx via Content-Controls (SDT).

build_brief(
    sender_oe, sender_first_name, sender_last_name,
    sender_street, sender_zip, sender_city,
    sender_contact_email,
    sender_contact_phone=None, sender_contact_mobile=None,
    recipient_salutation, recipient_first_name, recipient_last_name,
    recipient_company=None, recipient_street, recipient_zip, recipient_city,
    date_city, date, subject, body,
    introduction=None, closing="Freundliche Grüsse",
    signatory_name=None, signatory_role=None,
    enclosures=None, copy_to=None,
    output_path="brief.docx"
)
"""
import io, os, zipfile
from datetime import date as _date
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_TEMPLATE = os.path.join(os.path.dirname(__file__), "assets", "Vorlage_Brief.dotx")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_MONTHS_DE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
              "Juli", "August", "September", "Oktober", "November", "Dezember"]

def _today_de():
    """Aktuelles Datum im Galledia-Format '1. Juni 2026'."""
    t = _date.today()
    return f"{t.day}. {_MONTHS_DE[t.month - 1]} {t.year}"


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
    """Findet alle Content Controls (SDT) im Dokument und retourniert dict[tag] = sdt_element."""
    sdts = {}
    body = doc.element.body
    # Iteriere über alle SDT-Elemente
    for sdt in body.iter(qn('w:sdt')):
        # Tag-Element finden
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


def _add_text_runs(paragraph, text, multiline=False, rPr_template=None):
    """Fügt Text als Runs (mit Line-Breaks) zu einem Paragraph hinzu.
    rPr_template wird in jeden Run kopiert (Font/Style übernehmen)."""
    if not text:
        paragraph.append(_make_run("", rPr_template))
        return
    lines = text.split('\n') if multiline else [text]
    for i, line in enumerate(lines):
        if i > 0:
            paragraph.append(_make_break_run(rPr_template))
        paragraph.append(_make_run(line, rPr_template))


def _replace_body_text(doc, body_text):
    """Findet den '[Text]'-Platzhalter-Paragraph zwischen Introduction und Closing,
    ersetzt ihn durch eine Folge von Paragraphen mit dem body_text.

    body_text: String mit \\n\\n als Absatz-Trenner, \\n als Zeilenumbruch innerhalb Absatz.
    """
    body = doc.element.body
    # Sammle alle Paragraphen die [Text] enthalten
    text_paragraphs = []
    for p in body.iter(qn('w:p')):
        # Check if this paragraph contains "[Text]" in any w:t
        para_text = "".join(t.text or "" for t in p.iter(qn('w:t')))
        if "[Text]" in para_text:
            text_paragraphs.append(p)

    if not text_paragraphs:
        return  # kein Platzhalter gefunden

    target_p = text_paragraphs[0]  # erster [Text]-Paragraph = Body-Slot

    # Style-Info vom Target-Paragraph übernehmen (pPr)
    target_pPr = target_p.find(qn('w:pPr'))

    # Body-Text in Absätze aufteilen
    absatz_texte = body_text.split('\n\n') if body_text else [""]

    # Position im Parent merken
    parent = target_p.getparent()
    target_idx = list(parent).index(target_p)

    # Neue Paragraphen erstellen — mit LEEREM Paragraph zwischen Absätzen für visuelle Leerzeile
    from copy import deepcopy
    new_paragraphs = []
    for i, absatz in enumerate(absatz_texte):
        # Vor jedem Absatz (ausser dem ersten) eine Leerzeile einfügen
        if i > 0:
            empty_p = OxmlElement('w:p')
            if target_pPr is not None:
                empty_p.append(deepcopy(target_pPr))
            # Leerer Paragraph braucht einen leeren Run, sonst kollabiert er
            _add_text_runs(empty_p, "")
            new_paragraphs.append(empty_p)
        # Inhalts-Paragraph
        new_p = OxmlElement('w:p')
        if target_pPr is not None:
            new_p.append(deepcopy(target_pPr))
        _add_text_runs(new_p, absatz, multiline=True)
        new_paragraphs.append(new_p)

    # Target-Paragraph durch neue Paragraphen ersetzen
    parent.remove(target_p)
    for i, np in enumerate(new_paragraphs):
        parent.insert(target_idx + i, np)


_OAW_NS = "http://schemas.officeatwork.com/CustomXMLPart"


def _patch_office_at_work(docx_path, values):
    """Patcht den officeatwork CustomXMLPart (customXml/itemN.xml) im
    gespeicherten docx, sodass Word die Werte beim Öffnen via <w:dataBinding>
    in alle SDTs schreibt. Ohne diesen Patch zeigt Word den Glossary-Default.

    values: dict[tag, text] — Tag-Namen wie Date, City, RecipientAddress.
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
    """Baut Empfänger-Adress-Block (mehrzeilig mit \\n)."""
    lines = []
    salut_str = f"{salutation} " if salutation else ""
    lines.append(f"{salut_str}{first_name} {last_name}".strip())
    if company:
        lines.append(company)
    lines.append(street)
    lines.append(f"{zip_code} {city}".strip())
    return "\n".join(lines)


def _sender_org_block(oe, street, zip_code, city):
    """Absender-Organisation-Block (oben links, klein)."""
    return f"{oe} · {street} · {zip_code} {city}"


def _doc_contact_block(first_name, last_name, email, phone=None, mobile=None):
    """Sachbearbeiter-Block (rechts unter Adressfeld)."""
    lines = [f"{first_name} {last_name}", email]
    if phone:
        lines.append(phone if phone.startswith("T ") else f"T {phone}")
    if mobile:
        lines.append(mobile if mobile.startswith("M ") else f"M {mobile}")
    return "\n".join(lines)


def _signature_block(name, role=None):
    """Unterschriften-Block (Name + optional Rolle)."""
    if role:
        return f"{name}\n{role}"
    return name


def build_brief(
    sender_oe,
    sender_first_name, sender_last_name,
    sender_street, sender_zip, sender_city,
    sender_contact_email,
    recipient_salutation, recipient_first_name, recipient_last_name,
    recipient_street, recipient_zip, recipient_city,
    subject, body,
    date_city=None, date=None,
    sender_contact_phone=None, sender_contact_mobile=None,
    recipient_company=None,
    introduction=None,
    closing="Freundliche Grüsse",
    signatory_name=None, signatory_role=None,
    signatory_name_2=None, signatory_role_2=None,
    enclosures=None, copy_to=None,
    template=None,
    output_path="brief.docx",
):
    """Erstellt einen Galledia-Geschäftsbrief aus Vorlage_Brief.dotx.

    Args:
        sender_*: Absender-Daten (Galledia-Mitarbeiter)
            sender_oe: OE-Schreibweise ("Galledia Fachmedien AG" etc.)
            sender_first_name, sender_last_name: Vor- und Nachname
            sender_street: "Buckhauserstrasse 24"
            sender_zip, sender_city: "8048", "Zürich"
            sender_contact_email: "stefan.zimmermann@galledia.ch"
            sender_contact_phone: "T +41 58 344 96 22" (optional, eines von beiden Pflicht)
            sender_contact_mobile: "M +41 79 555 12 34" (optional)
        recipient_*: Empfänger-Daten
            recipient_salutation: "Herr" / "Frau" (mit ggf. Titel: "Frau Dr.")
            recipient_first_name, recipient_last_name: Vor- und Nachname
            recipient_company: optional, Firma
            recipient_street, recipient_zip, recipient_city: Adresse
        date_city: Absendeort ("Zürich")
        date: Datum ("1. Juni 2026")
        subject: Betreff
        body: Brieftext. \\n\\n = Absatztrenner, \\n = Zeilenumbruch innerhalb Absatz.
        introduction: Anrede. Default: "Sehr geehrte/r Herr/Frau {Nachname}" basierend auf salutation
                      Falls Firma ohne Person: "Sehr geehrte Damen und Herren"
        closing: Gruss, Default "Freundliche Grüsse"
        signatory_name: Default = sender_first_name + sender_last_name
        signatory_role: optional, Rolle des Unterschreibenden
        signatory_name_2 / signatory_role_2: zweite Unterschrift (optional)
        enclosures: optional, list[str] der Beilagen
        copy_to: optional, list[str] CC-Empfänger
        template: optionaler alternativer Pfad zur Vorlage
        output_path: Output-Pfad für die .docx

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

    # Default-Anrede berechnen falls nicht gesetzt
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

    # AdressBlock-Komposition entspricht dem officeatwork-Original:
    # - Organisation       = nur Firmen-Name           ("Galledia Fachmedien AG")
    # - Address1           = Strasse                   ("Baslerstrasse 60")
    # - Address2           = PLZ + Ort                 ("8048 Zürich")
    # - AdressBlock        = Strasse + PLZ Ort, mehrzeilig
    # - AdressBlockOrganisation = nur Firmen-Name
    # - Doc.Contact        = LABEL "Ihr Kontakt" (KEIN Block!)
    # - AdressBlockSignature1 = voller Kontaktblock (Name + T + M + Email)
    # - Signature1         = Name (+ optional Rolle)
    contact_block = _doc_contact_block(
        sender_first_name, sender_last_name, sender_contact_email,
        sender_contact_phone, sender_contact_mobile)
    sig1 = _signature_block(signatory_name, signatory_role)

    # Alle Felder dynamisch aus Args. AdressBlock (Sender, oberhalb Empfänger)
    # einzeilig: "Strasse · PLZ Ort" — multi-line würde mit Empfänger überlappen.
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
    }
    if signatory_name_2:
        sig2 = _signature_block(signatory_name_2, signatory_role_2)
        mappings["Signature2"] = sig2
        mappings["AdressBlockSignature2"] = sig2

    if copy_to:
        if isinstance(copy_to, list):
            mappings["CopyTo"] = "Kopie an: " + ", ".join(copy_to)
        else:
            mappings["CopyTo"] = f"Kopie an: {copy_to}"

    multiline_keys = {
        "AdressBlock", "AdressBlockSignature1", "AdressBlockSignature2",
        "RecipientAddress", "Signature1", "Signature2",
    }
    # Belt-and-suspenders: setze SDT-Content UND patche CustomXMLPart
    for tag, val in mappings.items():
        if tag in sdts:
            _set_sdt_text(sdts[tag], val, multiline=(tag in multiline_keys))

    # Betreff: oft als Titel-Style-Paragraph oder in einem SDT — versuche beide
    # In dieser Vorlage gibt es keinen "Subject"-SDT — wir prüfen ob ein Paragraph
    # im Titel-Style vorhanden ist und ersetzen den Inhalt
    if subject:
        for p in doc.element.body.iter(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None and pStyle.get(qn('w:val')) == "Titel":
                    # Existierende Runs leeren und Subject einfügen
                    for r in p.findall(qn('w:r')):
                        p.remove(r)
                    _add_text_runs(p, subject)
                    break  # nur ersten Titel-Paragraph

    # Body-Text einfügen (zwischen Introduction und Closing)
    _replace_body_text(doc, body)

    # Beilagen (optional, als zusätzlicher Paragraph am Ende vor signature/copy_to)
    # Skip falls keine Beilagen; der Code-Workflow ist optional
    if enclosures:
        encl_text = "Beilagen: " + ", ".join(enclosures) if isinstance(enclosures, list) else f"Beilagen: {enclosures}"
        # Füge nach dem Closing-Paragraph einen neuen Paragraph mit den Beilagen hinzu
        body_el = doc.element.body
        for sdt in body_el.iter(qn('w:sdt')):
            sdtPr = sdt.find(qn('w:sdtPr'))
            if sdtPr is None:
                continue
            tag_el = sdtPr.find(qn('w:tag'))
            if tag_el is None or tag_el.get(qn('w:val')) != "Closing":
                continue
            # Closing-SDT gefunden — finde dessen Eltern-Paragraph
            closing_p = sdt
            while closing_p.tag != qn('w:p') and closing_p.getparent() is not None:
                closing_p = closing_p.getparent()
            if closing_p.tag == qn('w:p'):
                parent = closing_p.getparent()
                idx = list(parent).index(closing_p)
                # Beilagen-Paragraph (mit Style "Absender" oder normal) nach Closing einfügen
                encl_p = OxmlElement('w:p')
                _add_text_runs(encl_p, encl_text)
                parent.insert(idx + 1, encl_p)
            break

    doc.save(output_path)
    # Final pass: patch officeatwork CustomXMLPart (Bug-Fix v2.0.1)
    # Word liest Werte via <w:dataBinding xpath="/ns:officeatwork/ns:Date">
    # aus dem CustomXMLPart — ohne diesen Patch zeigt Word den Glossary-Default
    _patch_office_at_work(output_path, mappings)
    return output_path
