---
name: galledia-dokument
description: "Erstelle CI-konforme Galledia-Dokumente (.docx): Angebote, Offerten, Dokumentationen, Schulungsunterlagen. Triggers: 'erstelle ein Dokument', 'Offerte schreiben', 'Angebot', 'Dokumentation', '.docx', 'Word-Dokument'. Nutzt Vorlage_Dokument.dotx mit Galledia-Header (G-Logo), Fusszeile (Titel + Seitenzahl), DokTitel/DokUntertitel/Überschriften 1–3, Bullets, Tabellen."
version: "0.0.6"
---

# Galledia Dokument Skill

CI-konforme Word-Dokumente aus `Vorlage_Dokument.dotx` via `fill_dokument.py`.

## Pflichtabfrage vor dem Start

Immer nachfragen, falls nicht im Gespräch bekannt:
- **Dokumenttitel** (z.B. «KI-Hub Offerte», «Schulungsunterlagen Mailchimp»)
- **Untertitel / Anlass** (z.B. «Angebot für XY AG»)
- **Datum** (Ort und Datum, z.B. «Stäfa, 29. Mai 2026»)
- **Rechtseinheit** (z.B. «Galledia Fachmedien AG»)
- **Adresse** (z.B. «Seestrasse 90a, 8712 Stäfa»)
- **Struktur** (Kapitel und Inhalte — wenn nicht klar, erst Gliederung vorschlagen und bestätigen lassen)

## Setup

**Wichtig:** Setup-Block ist Python und MUSS als `python3 - <<'EOF' ... EOF`
Heredoc ausgeführt werden — sonst interpretiert Bash `import` als ImageMagick.

```bash
pip install python-docx --break-system-packages
```

```bash
python3 - <<'EOF'
import os, sys, urllib.request
_DIR = "/tmp/galledia_dokument"
os.makedirs(f"{_DIR}/assets", exist_ok=True)
sys.path.insert(0, _DIR)
_BASE = "https://raw.githubusercontent.com/galledia-ag/galledia-office-ci-assets/main/skills/galledia-dokument"
for _name in ["fill_dokument.py", "assets/Vorlage_Dokument.dotx"]:
    _dest = f"{_DIR}/{_name}"
    if not os.path.exists(_dest):
        try:
            urllib.request.urlretrieve(f"{_BASE}/{_name}", _dest)
            print(f"✓ {_name} ({os.path.getsize(_dest):,} bytes)")
        except Exception as e:
            raise RuntimeError(
                f"Asset-Download fehlgeschlagen ({_name}): {e}. "
                f"Plugin/Repo nicht erreichbar — Skill abbrechen, NICHT mit "
                f"python-docx from scratch improvisieren."
            )
print("Assets bereit")
EOF
```

Der eigentliche Build-Aufruf (`from fill_dokument import build_document`)
erfolgt im selben Code-Execution-Kontext in einem zweiten `python3 -` Heredoc.

## Verwendung

```python
build_document(
    titel       = "KI-Hub Offerte",
    untertitel  = "Angebot AI-Infrastruktur für Fachmedien",
    datum       = "Stäfa, 29. Mai 2026",
    rechtseinheit = "Galledia Fachmedien AG",
    adresse     = "Seestrasse 90a\n8712 Stäfa",   # \n = Zeilenumbruch
    empfaenger  = "XY AG\nz.H. Herr Max Muster\nMusterstrasse 1\nCH-8000 Zürich",  # optional, \n = Zeilenumbruch
    abschnitte  = [
        {
            "titel": "Ausgangslage",          # → Überschrift 1 (auto-nummeriert)
            "inhalt": [
                {"typ": "text",    "inhalt": "Fliesstext..."},
                {"typ": "bullet",  "inhalt": "Aufzählungspunkt"},
                {"typ": "h2",      "inhalt": "Unterkapitel"},    # → Überschrift 2
                {"typ": "h3",      "inhalt": "Abschnitt"},       # → Überschrift 3
                {"typ": "tabelle", "inhalt": [
                    ["Spalte A", "Spalte B"],   # Erste Zeile = Kopfzeile (rot)
                    ["Wert 1",   "Wert 2"],
                ]},
            ]
        },
    ],
    output_path = "output.docx",
)
```

## Ausgabe-Layout

| Element | Style | Formatierung |
|---|---|---|
| Dokumenttitel | DokTitel | Volte Semibold, 36 pt |
| Untertitel | DokUntertitel | 22 pt |
| Datum / Ort | DokDatum | Vorlage |
| Absender | DokAbsender | Rechtseinheit / Adresse |
| Empfänger | DokEmpfaenger | Mehrzeilig via \n |
| Kapitel | Überschrift 1 | Volte Semibold, auto-nummeriert 1. 2. 3. |
| Unterkapitel | Überschrift 2 | auto-nummeriert 1.1 1.2 |
| Abschnitt | Überschrift 3 | |
| Fliesstext | Standard | Volte Regular, 11 pt |
| Aufzählung | DokBullet | Punkt |
| Tabelle Kopf | Tabellen-Titel | Volte Semibold, Rot #E61C52, Hintergrund #F2F2F5 |
| Tabelle Inhalt | Tabellen-Inhalt | Volte Regular |
| Inhaltsverzeichnis | automatisch | In Word mit F9 aktualisieren |
| Kopfzeile | Header | Galledia G-Logo rechts |
| Fusszeile | Footer | Dokumenttitel links, Seite N von M rechts |

## CI-Regeln

- Keine Versalien in Überschriften
- Schweizer Rechtschreibung (kein ß)
- Tabellen immer mit Kopfzeile (erste Zeile)
- Fliesstext nach Tabellen: `{'typ': 'text', 'inhalt': ''}` für Abstand
- Empfänger-Block nur wenn vorhanden (optional)
- Inhaltsverzeichnis: in Word mit F9 aktualisieren oder Strg+A → F9

## Aufbau

Deckblatt → TOC (eigene Seite) → Inhalt. Ort/Datum und Rechtseinheit stehen nur auf dem Deckblatt.

## Hinweis Inhaltsverzeichnis

Das TOC-Feld wird korrekt generiert. Word aktualisiert es beim ersten Öffnen automatisch oder manuell via F9. In LibreOffice: Extras → Felder aktualisieren.
