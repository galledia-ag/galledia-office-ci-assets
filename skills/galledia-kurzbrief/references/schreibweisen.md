# Schreibweisen — Galledia CI/CD (Markenhandbuch v1.5)

## Organisationseinheiten (5 — exakte Schreibweise)

| Schreibweise | Typ |
|---|---|
| `galledia group ag` | Dachmarke / Holding — **komplett klein** |
| `Galledia Fachmedien AG` | Verlag (Frauenfeld, Luzern, Zuerich) |
| `Galledia Regionalmedien AG` | Regionalmedien (Berneck) |
| `Galledia Print AG` | Druckerei |
| `Galledia Digital AG` | Digital |

### Verboten — diese Schreibweisen NIE verwenden

- ~~Galledia AG~~ (gibt es nicht)
- ~~Galledia Gruppe~~ / ~~Galledia Group AG~~ (Holding heisst klein "galledia group ag")
- ~~Galledia GmbH~~ (alles AG)
- ~~galledia fachmedien ag~~ / ~~galledia print ag~~ (nur die Holding wird klein geschrieben)
- ~~GALLEDIA FACHMEDIEN AG~~ (keine Versalien ausser bei Eigennamen wie `SAVE AG`)

## Co-Branding-Zusatz

Bei Publikationen / Zeitschriften:
- `Eine Publikation der Galledia.` (Einzahl)
- `Publikationen der Galledia` (Mehrzahl)

Bei Monomarken-Websites:
- `Ein Kompetenzzentrum der Galledia Print AG` (Footer)

## Telefon / Mobile / Fax

| Format | Beispiel |
|---|---|
| Festnetz | `T +41 58 344 96 22` |
| Mobile | `M +41 78 846 24 16` |
| Trenner (in einer Zeile) | `T +41 58 344 96 22 \| M +41 78 846 24 16` |

**Regeln:**
- Immer Laendervorwahl `+41` (auch fuer CH-Nummern)
- Praefix `T` (Telefon) oder `M` (Mobile) — nicht `Tel.`, nicht `Tel`, nicht `Phone`
- Leerzeichen zwischen Blocks: `+41 58 344 96 22` (Vorwahl Block + 3-2-2)
- **Fax wird NICHT mehr verwendet** — niemals in Briefe aufnehmen

## Sonderzeichen

| Zweck | Zeichen | Unicode | NICHT verwenden |
|---|---|---|---|
| Aufzaehlung | `·` (Mittelpunkt) | U+00B7 | `-`, `•`, `*`, `–` |
| Anfuehrungszeichen | `«` `»` (Guillemets) | U+00AB / U+00BB | `"`, `"`, `"`, `\`, `'` |
| Trenner (Zeile / Adresse) | `\|` (Pipe) | U+007C | `/`, `-`, `,` |
| Gedankenstrich | `—` (em dash) | U+2014 | `-`, `--` |

## Schreibweise-Pruefung (Regex fuer Validierung)

```python
VALID_OE = [
    "galledia group ag",
    "Galledia Fachmedien AG",
    "Galledia Regionalmedien AG",
    "Galledia Print AG",
    "Galledia Digital AG",
]

PHONE_RE = r"^[TM] \+41 \d{2,3}( \d{2,3}){2,3}$"
# matches: T +41 58 344 96 22, M +41 78 846 24 16, T +41 71 757 75 85

FORBIDDEN_PATTERNS = [
    (r"\bGalledia AG\b", "Es gibt keine 'Galledia AG' — exakte OE verwenden"),
    (r"\bGalledia Gruppe\b", "Heisst 'galledia group ag' (klein)"),
    (r"\bGalledia Group AG\b", "Heisst 'galledia group ag' (klein)"),
    (r"\bGalledia GmbH\b", "Alle Galledia-OE sind AG"),
    (r"\bFax\b", "Fax wird nicht mehr verwendet"),
]
```

## Quelle

Galledia Markenhandbuch v1.5 (7.6.2021), insbesondere S. 5 (Kurzfassung), S. 30 (Schrift) und S. 43 (Adressen).
