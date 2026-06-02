---
name: galledia-kurzbrief
description: >
  Erstellt einen Galledia-Kurzbrief (transmittal sheet / kurze Mitteilung)
  im CI/CD gemaess Markenhandbuch v1.5. Verwende diesen Skill immer wenn
  ein User einen Kurzbrief, eine kurze Mitteilung, ein Begleitschreiben
  fuer Unterlagen, eine Aktennotiz oder ein Memo mit Standardoptionen
  ("zur Kenntnisnahme", "zur Erledigung", "Beilagen" etc.) erstellen will.
  Triggert auf: "Kurzbrief", "Kurzmitteilung", "kurze Mitteilung",
  "Begleitschreiben", "Aktennotiz", "Memo", "Transmittal", "Beilage zur".
  Im Gegensatz zum normalen Brief (galledia-brief) hat der Kurzbrief
  KEINEN Brieftext-Body, sondern 10 vordefinierte Notizoptionen mit
  Checkboxen. Generierung ueber Code Execution (fill_kurzbrief.py +
  Vorlage_Kurzbrief.dotx) direkt im Sandbox.
  Arbeitssprache: Schweizer Hochdeutsch.
version: "2.0.4"
---

# Galledia Kurzbrief (v2.0.4 — Code Execution)

1-seitiges Begleitschreiben fuer Beilagen mit 10 vordefinierten
Notizoptionen (Checkboxen — werden automatisch mit ☒ angekreuzt
basierend auf den im Build-Aufruf angegebenen Notes).

Generierung direkt im Sandbox aus `Vorlage_Kurzbrief.dotx` via `fill_kurzbrief.py`.
Funktioniert in claude.ai Web, Cowork und Claude Desktop.

## Workflow

### Schritt 1 — Daten sammeln (Hard-Stop bei fehlenden Pflichtfeldern)

**Kurzbriefe muessen versandfertig generiert werden — keine Platzhalter.**
Wenn ein Pflichtfeld fehlt ODER Claude sich unsicher ist: **STOP**, gezielte
Rueckfrage, KEIN Build bevor alles vollstaendig ist.

#### Pflichtfelder — ABSENDER (Galledia-Mitarbeiter)

| Feld | Beispiel | Pflicht |
|---|---|---|
| `sender_oe` (Aktiengesellschaft) | `Galledia Fachmedien AG` | ✅ |
| `sender_first_name` | `Stefan` | ✅ |
| `sender_last_name` | `Zimmermann` | ✅ |
| `sender_street` | `Buckhauserstrasse 24` | ✅ |
| `sender_zip` + `sender_city` | `8048` + `Zürich` | ✅ |
| `sender_contact_email` | `stefan.zimmermann@galledia.ch` | ✅ |
| `sender_contact_phone` ODER `sender_contact_mobile` | `T +41 58 ...` / `M +41 79 ...` | ✅ (eines) |
| `signatory_role` | `Leitung Fachmedien & Digital \| Mitglied der Gruppenleitung` | ✅ (Funktion unter dem Namen) |

OE: eine der 5 exakten Schreibweisen (klein für galledia group ag).
`signatory_role` zwingend angeben — sonst fehlt die Funktion unter dem Namen.
Bei Bedarf aus User-Memory ziehen.

#### Pflichtfelder — EMPFAENGER

| Feld | Beispiel | Pflicht |
|---|---|---|
| `recipient_salutation` | `Herr` / `Frau` | ✅ |
| `recipient_first_name` | `Noëlla` | ✅ |
| `recipient_last_name` | `Schüttel` | ✅ |
| `recipient_company` | `Felderhus AG` | ⬜ optional |
| `recipient_street` | `Neuhofstrasse 26` | ✅ |
| `recipient_zip` + `recipient_city` | `6345` + `Neuheim` | ✅ |

#### Pflichtfelder — ALLGEMEIN

| Feld | Auto-Default | Beispiel |
|---|---|---|
| `date_city` | ← `sender_city` falls leer | `Zuerich` |
| `date` | ← heutiges Datum | `1. Juni 2026` |
| `subject` | (Pflicht) | `Felderhus` |

Kein `body` — stattdessen werden die `notes` (siehe unten) angekreuzt.

#### Anrede — HART (nie generisch wenn Name bekannt)

| Empfaenger | `introduction` (auto) |
|---|---|
| `Herr`, Name bekannt | `Sehr geehrter Herr {Nachname}` |
| `Frau`, Name bekannt | `Sehr geehrte Frau {Nachname}` |
| Nur Firmenadresse | `Sehr geehrte Damen und Herren` |

### Notes — die 10 Standardoptionen

Aus dem Markenhandbuch v1.5, fest in der Vorlage verankert (Checkboxen):

| Key | Default-Text |
|---|---|
| `Note1` | zur Kenntnisnahme |
| `Note2` | zu Ihren Akten |
| `Note3` | auf Ihren Wunsch |
| `Note4` | mit Dank zurück |
| `Note5` | zur Erledigung |
| `Note6` | gemäss telefonischer Besprechung |
| `Note7` | zur Stellungnahme |
| `Note8` | gemäss Ihrer Anfrage |
| `Note9` | per E-Mail an: |
| `Note10` | Beilagen: |

**Ankreuzen — zwei Wege:**
```python
# (1) Mit Text-Override → wird AUTOMATISCH angekreuzt (☒)
notes = {"Note10": "Beilagen: Vertrag, AGB, NDA"}

# (2) Nur ankreuzen, Default-Text behalten
checked = ["Note1", "Note5"]  # ☒ "zur Kenntnisnahme", ☒ "zur Erledigung"
```

**Wichtig:** Jedes Note-Key in `notes` UND jedes Note-Key in `checked`
wird in der finalen .docx mit ☒ angekreuzt. Notes nicht in einer der
beiden Listen bleiben ☐ (leer) und tragen den Default-Text.

Note9 / Note10 sind die haeufigsten Ueberschreib-Kandidaten (E-Mail-Verteiler,
konkrete Beilagen-Liste).

### Schritt 2 — Versandfertig-Checkliste (Pflicht vor Build)

```
Bereit zum Generieren — bitte kurz pruefen:

ABSENDER:    Stefan Zimmermann, Galledia Fachmedien AG
             Buckhauserstrasse 24, 8048 Zuerich
             stefan.zimmermann@galledia.ch, T +41 58 344 96 22

EMPFAENGER:  Frau Noëlla Schüttel
             Neuhofstrasse 26, 6345 Neuheim

DATUM:       Zuerich, 1. Juni 2026
BETREFF:     Felderhus
ANREDE:      Sehr geehrte Frau Schüttel
ANGEKREUZT:  Beilagen: gemäss Besprechung
GRUSS:       Freundliche Gruesse / Stefan Zimmermann

Alles korrekt?
```

### Schritt 3 — Setup + Build via Code Execution

```bash
pip install python-docx --break-system-packages
```

```python
import os, sys, urllib.request
_DIR = "/tmp/galledia_kurzbrief"
os.makedirs(f"{_DIR}/assets", exist_ok=True)
sys.path.insert(0, _DIR)
_BASE = "https://raw.githubusercontent.com/galledia-ag/galledia-office-ci-assets/main/skills/galledia-kurzbrief"
for _name in ["fill_kurzbrief.py", "assets/Vorlage_Kurzbrief.dotx"]:
    _dest = f"{_DIR}/{_name}"
    if not os.path.exists(_dest):
        try:
            urllib.request.urlretrieve(f"{_BASE}/{_name}", _dest)
        except Exception as e:
            raise RuntimeError(
                f"Asset-Download fehlgeschlagen ({_name}): {e}. "
                f"Plugin/Repo nicht erreichbar — Skill abbrechen, NICHT improvisieren."
            )

from fill_kurzbrief import build_kurzbrief

out = build_kurzbrief(
    sender_oe="Galledia Fachmedien AG",
    sender_first_name="Stefan", sender_last_name="Zimmermann",
    sender_street="Buckhauserstrasse 24", sender_zip="8048", sender_city="Zürich",
    sender_contact_email="stefan.zimmermann@galledia.ch",
    sender_contact_phone="T +41 58 344 96 22",
    recipient_salutation="Frau",
    recipient_first_name="Noëlla", recipient_last_name="Schüttel",
    recipient_street="Neuhofstrasse 26", recipient_zip="6345", recipient_city="Neuheim",
    subject="Felderhus",
    notes={"Note10": "Beilagen: gemäss Besprechung"},
    # date_city, date werden auto-gesetzt
    output_path="/home/claude/Kurzbrief_Schüttel.docx",
)
print(f"Kurzbrief erstellt: {out}")
```

### Schritt 4 — User die fertige .docx liefern

Datei mit `present_files` an den User.

## CI-Regeln

Identisch zum Brief — siehe `references/schreibweisen.md`.

## Was NICHT in diesem Skill

- Voller Brief mit Brieftext → `galledia-brief`
- Praesentation → `galledia-praesentation`
- Mehrseitiges Dokument → `galledia-dokument`
