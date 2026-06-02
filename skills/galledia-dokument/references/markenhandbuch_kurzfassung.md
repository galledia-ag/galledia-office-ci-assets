# Galledia Markenhandbuch — Kurzfassung (v1.5)

Extrahiert aus `GA_Markenhandbuch-low.pdf` (7.6.2021). Nur die fuer die Brief-Generierung relevanten Teile.

## Logo

- Wort-Bild-Marke aus "G"-Bildmarke und Wortmarke "Galledia"
- Farben: Galledia-Rot, Weiss oder Galledia-Schwarz (nie andere)
- Schutzzone: ergibt sich aus den Abstaenden Bild- zu Wortmarke (in Brief-Vorlage bereits korrekt)
- **Logo ohne Schriftzug** nur, wenn unmittelbar darunter ein Adressblock folgt (= Brief, Visitenkarte)
- **Don'ts**: Logo mit Umform, verzerrtes Logo, anderes Verhaeltnis Wort-/Bildmarke

In der Brief-Vorlage ist das Logo bereits korrekt platziert — **nicht antasten**.

## Farben (Primaer)

| Name | Pantone | CMYK | RGB | HEX |
|---|---|---|---|---|
| Galledia-Rot | 206 U | 0/95/50/0 | 230/28/82 | `#E61C52` |
| Galledia-Schwarz | — | 0/0/0/100 | 0/0/0 | `#000000` |

Hinweis: Im Handbuch ist HEX teils mit Druckfehler als `#E17F9C` angegeben — der korrekte Wert fuer RGB 230/28/82 ist `#E61C52`.

## Farben (Sekundaer + Zusatz)

Galledia-Grau 1–4 (#404040, #666666, #A6A6A6, #D9D9D9) sowie Bronze (#B27547), Purple (#9E3F6B), Blau (#4A52A5), Tuerkis (#22AA9F). Im Brief in der Regel nicht relevant — Standard ist Schwarz auf Weiss mit Rot-Logo.

## Typografie

| Schrift | Anwendung |
|---|---|
| Volte Rounded Regular / Semibold | Plakative Titel (Titelseiten — im Brief nicht relevant) |
| Volte Regular | Grundtext, Fliesstext (Brief-Standard, 10pt, ZA 12.5pt) |
| Volte Regular Italic | Hervorhebungen im Text |
| Volte Semibold | Zwischentitel, Auszeichnungen |

**Brief-Standard**: Volte Regular, 10pt, Zeilenabstand 12.5pt, linksbuendig, Schwarz.

**Verfuegbarkeit**: Auf allen 200 Galledia-Geraeten installiert. Web-Fallback: `.woff` von galledia.ch.

## Sonderzeichen

- Aufzaehlung: `·` (Mittelpunkt)
- Anfuehrungszeichen: `«` `»` (Guillemets)
- Adress-Trenner: `|`

## Elemente

- Kasten / Linien: **abgerundete Ecken** (in Brief-Vorlage bereits so)

## Was im Brief NICHT vorkommt (Markenhandbuch beachtet, aber Brief-irrelevant)

- Piktogramme (S. 29) — nur in Image-/Produkt-Kommunikation
- Alphabet (S. 31) — nur fuer plakative Titel
- Bildwelt (S. 33-39) — nicht im Geschaeftsbrief

## Update-Prozess

Aenderungen am Markenhandbuch fuehren zu PR im Plugin-Repo:
- `references/*` aktualisieren
- ggf. `templates/Brief-Vorlage Galledia.dotx` ersetzen (durch Designer:in via PR, nicht editiert)
- Version in `.claude-plugin/plugin.json` bumpen
