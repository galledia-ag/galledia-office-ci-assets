---
name: galledia-brief
description: >
  Erstellt einen Geschaeftsbrief im Galledia-CI/CD (Markenhandbuch v1.5).
  Verwende diesen Skill IMMER, wenn der User einen Brief, Geschaeftsbrief,
  Anschreiben, Kundenbrief, Offertschreiben, Mahnung oder ein aehnliches
  Schreiben im Namen einer Galledia-Organisationseinheit erstellen will.
  Triggert auf: "Brief", "schreibe einen Brief", "Geschaeftsbrief",
  "Anschreiben", "Kundenbrief", "Schreiben an", "Offertschreiben",
  "Begleitbrief", "Bewerbungsschreiben (von Galledia an Kandidat)".
  Liefert eine fertige .docx im Galledia-CI als Download.
  Arbeitssprache: Schweizer Hochdeutsch (ss statt scharf-s, also "Grüsse"
  nicht "Grüße"). Die Dokumentengenerierung erfolgt ueber Code Execution
  (fill_brief.py + Vorlage_Brief.dotx) direkt im Sandbox.
version: "2.0.5"
---

# Galledia Geschaeftsbrief (v2.0.5 — Code Execution)

CI-konformer Brief direkt im Sandbox aus `Vorlage_Brief.dotx` via `fill_brief.py`.
Funktioniert in claude.ai Web, Cowork und Claude Desktop.

## Workflow

### Schritt 1 — Daten sammeln (Hard-Stop bei fehlenden Pflichtfeldern)

**Briefe muessen versandfertig generiert werden — keine Platzhalter, keine
`<TODO>`-Markierungen.** Bei fehlendem Pflichtfeld ODER Unsicherheit: **STOP**,
gezielte Rueckfrage, KEIN Build bevor alles vollstaendig ist.

Frage NIE nach Daten, die der User schon genannt hat. Werte, die eindeutig
ableitbar sind (heutiges Datum, Sender-Stadt aus OE-Adresse): selbstaendig
setzen.

#### Pflichtfelder — ABSENDER (Galledia-Mitarbeiter)

| Feld | Beispiel | Pflicht |
|---|---|---|
| `sender_oe` (Aktiengesellschaft) | `Galledia Fachmedien AG` | ✅ |
| `sender_first_name` | `Stefan` | ✅ |
| `sender_last_name` | `Zimmermann` | ✅ |
| `sender_street` | `Buckhauserstrasse 24` | ✅ |
| `sender_zip` + `sender_city` | `8048` + `Zürich` | ✅ |
| `sender_contact_email` | `stefan.zimmermann@galledia.ch` | ✅ |
| `sender_contact_phone` ODER `sender_contact_mobile` (mind. eines) | `T +41 58 344 96 22` / `M +41 79 555 12 34` | ✅ (eines) |
| `signatory_role` | `Leitung Fachmedien & Digital \| Mitglied der Gruppenleitung` | ✅ (Funktion unter dem Namen) |

OE muss exakt eine der 5 sein: `galledia group ag` (klein!), `Galledia Fachmedien AG`,
`Galledia Regionalmedien AG`, `Galledia Print AG`, `Galledia Digital AG`.
`signatory_role` zwingend angeben — sonst fehlt die Funktion unter dem Namen.
Bei Bedarf aus User-Memory ziehen.

#### Pflichtfelder — EMPFAENGER

| Feld | Beispiel | Pflicht |
|---|---|---|
| `recipient_salutation` | `Herr` / `Frau` | ✅ |
| `recipient_first_name` | `Hans` | ✅ |
| `recipient_last_name` | `Mueller` | ✅ |
| `recipient_company` | `Mueller AG` | ⬜ optional |
| `recipient_street` | `Bahnhofstrasse 1` | ✅ |
| `recipient_zip` + `recipient_city` | `8001` + `Zuerich` | ✅ |

#### Pflichtfelder — ALLGEMEIN

| Feld | Auto-Default | Beispiel |
|---|---|---|
| `date_city` | ← `sender_city` falls leer | `Zuerich` |
| `date` | ← heutiges Datum auf Deutsch | `1. Juni 2026` |
| `subject` | (Pflicht) | `Offerte fuer Inserat Q3 2026` |
| `body` | (Pflicht) | siehe Body-Format unten |

#### Anrede — HART (nie generisch wenn Name bekannt)

| Empfaenger | `introduction` (auto) |
|---|---|
| `Herr`, Name bekannt | `Sehr geehrter Herr {Nachname}` |
| `Frau`, Name bekannt | `Sehr geehrte Frau {Nachname}` |
| Nur Firmenadresse | `Sehr geehrte Damen und Herren` |

**Niemals** `Sehr geehrte Damen und Herren` wenn Empfaenger-Name bekannt.

#### Body-Format

- Doppelte Newlines `\n\n` = Absatzwechsel (Leerzeile sichtbar)
- Einfache Newlines `\n` = Zeilenumbruch innerhalb Absatz

### Schritt 2 — Versandfertig-Checkliste (Pflicht vor Build)

Bevor `build_brief()` aufgerufen wird, fasse den Brief zusammen und hole vom User
eine Bestaetigung:

```
Bereit zum Generieren — bitte kurz pruefen:

ABSENDER:    Stefan Zimmermann, Galledia Fachmedien AG
             Buckhauserstrasse 24, 8048 Zuerich
             stefan.zimmermann@galledia.ch, T +41 58 344 96 22

EMPFAENGER:  Herr Hans Mueller
             Mueller AG, Bahnhofstrasse 1, 8001 Zuerich

DATUM:       Zuerich, 1. Juni 2026
BETREFF:     Offerte fuer Inserat Q3 2026
ANREDE:      Sehr geehrter Herr Mueller
GRUSS:       Freundliche Gruesse / Stefan Zimmermann
BEILAGEN:    Offerte_Q3_2026.pdf

Alles korrekt?
```

Bei "ja" → Build. Bei Korrekturen → anpassen.
Bei vager Antwort → nicht generieren, nachfragen.

### Schritt 3 — Setup + Build via Code Execution

**Wichtig:** Beide nachfolgenden Code-Blöcke sind Python und MÜSSEN als
`python3 - <<'EOF' ... EOF` Heredoc ausgeführt werden — sonst interpretiert
Bash `import` als ImageMagick-Tool und scheitert.

```bash
pip install python-docx --break-system-packages
```

```bash
python3 - <<'EOF'
import os, sys, urllib.request
_DIR = "/tmp/galledia_brief"
os.makedirs(f"{_DIR}/assets", exist_ok=True)
sys.path.insert(0, _DIR)
_BASE = "https://raw.githubusercontent.com/galledia-ag/galledia-office-ci-assets/main/skills/galledia-brief"
for _name in ["fill_brief.py", "assets/Vorlage_Brief.dotx"]:
    _dest = f"{_DIR}/{_name}"
    if not os.path.exists(_dest):
        try:
            urllib.request.urlretrieve(f"{_BASE}/{_name}", _dest)
        except Exception as e:
            raise RuntimeError(
                f"Asset-Download fehlgeschlagen ({_name}): {e}. "
                f"Plugin/Repo nicht erreichbar — Skill abbrechen, NICHT improvisieren."
            )
print("Assets bereit")
EOF
```

```bash
python3 - <<'EOF'
import sys
sys.path.insert(0, "/tmp/galledia_brief")
from fill_brief import build_brief

out = build_brief(
    sender_oe="Galledia Fachmedien AG",
    sender_first_name="Stefan", sender_last_name="Zimmermann",
    sender_street="Buckhauserstrasse 24", sender_zip="8048", sender_city="Zürich",
    sender_contact_email="stefan.zimmermann@galledia.ch",
    sender_contact_phone="T +41 58 344 96 22",
    signatory_role="Leitung Fachmedien & Digital | Mitglied der Gruppenleitung",
    recipient_salutation="Herr",
    recipient_first_name="Hans", recipient_last_name="Müller",
    recipient_company="Müller AG",
    recipient_street="Bahnhofstrasse 1", recipient_zip="8001", recipient_city="Zürich",
    subject="Offerte für Inserat Q3 2026",
    body="""Vielen Dank für unser angenehmes Gespräch letzte Woche.

Wie besprochen sende ich Ihnen anbei unsere Offerte für die Werbeplatzierung
im Q3 2026.

Wir freuen uns auf Ihre Rückmeldung bis 30. Juni 2026.""",
    enclosures=["Offerte_Q3_2026.pdf", "AGB.pdf"],
    output_path="/home/claude/Brief_Mueller.docx",
)
print(f"Brief erstellt: {out}")
EOF
```

### Schritt 4 — User die fertige .docx liefern

Datei mit `present_files` an den User. Kurze Bestaetigung welche Daten verwendet
wurden, keine Brieftext-Wiederholung im Chat.

## CI-Regeln (Markenhandbuch v1.5)

- OE exakt eine der 5 erlaubten Schreibweisen
- Telefonformat zwingend `T +41 58 ...` / `M +41 79 ...`
- Anführungszeichen: `« »` (Guillemets), keine geraden `"`
- Keine verbotenen Begriffe: "Galledia AG", "Galledia Gruppe", "Galledia GmbH", "Fax"
- Schweizer Hochdeutsch: ss statt ß

Siehe `references/schreibweisen.md`, `references/adressbloecke.md` und
`references/markenhandbuch_kurzfassung.md` fuer Details.

## Was NICHT in diesem Skill

- Kurzbrief mit Standardoptionen → `galledia-kurzbrief`
- PowerPoint-Praesentation → `galledia-praesentation`
- Mehrseitiges Dokument (Offerte/Anleitung) → `galledia-dokument`
