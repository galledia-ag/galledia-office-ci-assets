---
name: galledia-praesentation
description: >
  Erstellt und bearbeitet PowerPoint-Präsentationen im Galledia Corporate Design.
  Verwenden wenn: ein Deck, Slides oder .pptx für Galledia / Galledia Fachmedien / ZSW
  erstellt, gefüllt oder überarbeitet werden soll. Liefert CI-Mechanik (Layouts, Farben,
  Schriften, Regeln) — nicht den Inhalt.
version: "2.0.2"
template: assets/Vorlage_6.pptx
---

# Galledia-Präsentation (v2.0.2)

Produktions-Skill für CI-konforme Galledia-Präsentationen mit MBB-Level-Layouts.
56 Funktionen: Basis-Layouts, Frameworks (SWOT/Porter/BCG/McKinsey 9-Box/Wardley/BMC/...),
Advanced Charts (Waterfall/Tornado/Sankey/Heatmap/...), Storyline (SCQA/Exec-Summary/...),
Quality-Gates (audit_deck/MECE-Check/Action-Title-Linter).


### Sprint 1 — verfügbare Funktionen

**Footnotes (Quellen-Anker im MBB-Stil):**

```python
from helpers import add_footnote_marker, footnote_block

# In KPI-Folie: Quelle pro Zahl referenzieren
kpi_shape = kpi_grid(prs, [("96 GB", "VRAM"), ("CHF 240k", "Investition")],
                      kicker="Hardware", headline="Pilot-Setup")
# Marker neben der ersten KPI-Box
add_footnote_marker(slide, kpi_shape, num=1, position="top-right")

# Footnote-Block am Folienfuss (sammelt alle Quellen)
footnote_block(slide, [
    (1, "Datasheet NVIDIA H100, März 2026"),
    (2, "Interne Schätzung Q1 2026, Inspire 925"),
    (3, "Vergleichsofferte Lenovo, 28.04.2026"),
])
# → erzeugt: "¹ Datasheet NVIDIA H100, März 2026    ² Interne Schätzung..."
```

**Chart-Annotations (Hervorhebungen, Aussage-Pfeile):**

```python
from helpers import add_callout, add_highlight_line, add_value_label
from pptx.util import Inches

# Callout-Box mit Pfeil — "hier kippt der Trend"
add_callout(s,
    x=Inches(14), y=Inches(2.5), w=Inches(4), h=Inches(1),
    text="Kipppunkt Q3 2026",
    arrow_to=(Inches(12.5), Inches(4.2)))  # Pfeilspitze zeigt zu Chart-Punkt

# Horizontale Zielwert-Linie über einem Chart
add_highlight_line(s,
    Inches(1), Inches(4.5), Inches(19), Inches(4.5),
    label="Zielwert 850k",
    dashed=True, color=RED)

# Wert-Annotation an Chart-Punkt
add_value_label(s, (Inches(12.4), Inches(3.8)), 23.4, fmt="+%.1f%%")
# → erzeugt "+23.4%" in Galledia-Rot
```

**Wann was nutzen:**
- `add_footnote_marker` + `footnote_block` → IMMER bei Datenfolien mit konkreten Zahlen. MBB-Decks haben pro Zahl eine Quelle.
- `add_callout` → wenn eine Aussage auf der Folie hervorgehoben werden muss ("Hier liegt unser Hebel", "Diese Zahl ist der Schlüssel")
- `add_highlight_line` → für Zielwerte, Benchmarks, Schwellen ("ab hier rentabel"), oder Vor-/Nachher-Trennung
- `add_value_label` → Prozent- oder Differenz-Annotation an Daten-Punkten

### Sprint 2 — Data Visualization

**Waterfall (EBIT-Brücke / Variance-Analyse):**
```python
waterfall(prs, [
    ("EBIT 2025",     2400, "absolute"),    # grau, Pivot
    ("Umsatz +",       380, "delta_pos"),   # türkis, positiv
    ("Kosten +",      -180, "delta_neg"),   # rot, negativ
    ("Synergien +",    120, "delta_pos"),
    ("EBIT 2026",     2720, "total"),       # dunkelgrau, Zwischensumme
], kicker="Ergebnis-Bridge", headline="EBIT-Treiber 2025–2026")
```

**2×2 Matrix (Priorisierung, Impact/Effort, BCG-light):**
```python
matrix_2x2(prs,
    x_axis=("Aufwand niedrig", "Aufwand hoch"),
    y_axis=("Impact niedrig", "Impact hoch"),
    items=[
        (0.2, 0.85, "Quick Win A"),
        (0.7, 0.90, "Strategisch B"),
        (0.25, 0.30, "Aufschieben C"),
        (0.80, 0.20, "Vermeiden D"),
    ],
    kicker="Priorisierung", headline="Aufwand vs. Impact")
```

**Marimekko (Markt × Anteil mit variabler Marktgrösse):**
```python
mekko(prs, segments=[
    {"name": "Print",   "size": 0.45, "shares": [("Galledia", 0.35), ("Wettbewerb A", 0.25), ("Rest", 0.40)]},
    {"name": "Digital", "size": 0.30, "shares": [("Galledia", 0.20), ("Wettbewerb A", 0.45), ("Rest", 0.35)]},
    {"name": "Event",   "size": 0.25, "shares": [("Galledia", 0.40), ("Wettbewerb A", 0.30), ("Rest", 0.30)]},
], kicker="Marktanteile 2026", headline="Galledia stärkster im Event")
```

**100% Stacked Bar (Anteils-Entwicklung über Zeit/Kategorien):**
```python
stacked_bar_pct(prs,
    categories=["2024", "2025", "2026e"],
    series=[
        ("Print",   [45, 38, 30]),
        ("Digital", [35, 42, 48]),
        ("Event",   [20, 20, 22]),
    ],
    kicker="Sparten-Mix", headline="Print-Anteil sinkt linear")
```

**Quadrant Map (Strategische Positionierung — Scatter mit Achsen):**
```python
quadrant_map(prs,
    items_xy=[
        (0.75, 0.80, "Galledia Fachmedien", True),  # is_us=True → RED, grösser
        (0.30, 0.65, "Wettbewerb A"),
        (0.55, 0.40, "Wettbewerb B"),
    ],
    axes_labels=("Reichweite", "Tiefe der Beziehung"),
    kicker="Wettbewerbs-Landschaft", headline="Galledia stark positioniert")
```

**Wann was nutzen:**
- `waterfall` → Variance-Analyse (EBIT, Cash, Marktanteil-Veränderung) — IMMER wenn Vorher → Treiber → Nachher gezeigt wird
- `matrix_2x2` → Priorisierung, Quadranten-Klassifikation (Impact/Effort, Quick-Win/Strategisch/Aufschieben/Vermeiden)
- `mekko` → Markt-Übersicht mit Marktgrössen-Visualisierung (Marktanteile UND Markt-Bedeutung)
- `stacked_bar_pct` → Anteils-Entwicklung über Zeit oder Kategorien (Sparten-Mix, Demographics)
- `quadrant_map` → Wettbewerbs-Positionierung, „wo stehen wir" auf 2 Achsen — leichteres Layout als matrix_2x2

### Sprint 3 — Echte Charts via matplotlib

**Setup-Erweiterung:** zusätzlich zu `python-pptx Pillow` jetzt auch `matplotlib`:
```bash
pip install python-pptx Pillow matplotlib --break-system-packages
```

Volte-Fonts werden automatisch aus `assets/fonts/` für matplotlib registriert. Galledia-CI-Farbpalette (RED → TURK → G1 → G2 → G3 → G4) ist als Default-Cycle gesetzt. Top/Right Spines weg, light grid lines.

**Bar-Chart (vertikal/horizontal):**
```python
chart_bar(prs, [
    ("Print", 1620), ("Digital", 980), ("Event", 450), ("Sonstige", 220),
], kicker="Umsatz 2026", headline="Sparten-Übersicht", y_label="CHF k")

# Horizontal mit value_fmt:
chart_bar(prs, data, horizontal=True, value_fmt="{:,.0f} h")
```

**Line-Chart mit Annotations:**
```python
chart_line(prs,
    x=["2022", "2023", "2024", "2025", "2026e"],
    series={"Galledia": [100, 105, 112, 118, 124], "Markt": [100, 102, 104, 107, 110]},
    annotations=[{"x": "2026e", "y": 124, "text": "+14 vs. Markt"}],
    kicker="Indexierte Entwicklung", headline="Galledia über Markt-Wachstum",
    y_label="Index (2022=100)")
```

**Stacked Bar (absolute Werte — NICHT Prozent, dafür `stacked_bar_pct`):**
```python
chart_stacked_bar(prs,
    categories=["2024", "2025", "2026e"],
    series=[("Print", [1620, 1450, 1280]),
            ("Digital", [620, 780, 980]),
            ("Event", [380, 410, 450])],
    kicker="Umsatz nach Sparte", headline="Digital kompensiert Print",
    y_label="CHF k")
```

**Histogram (Verteilung):**
```python
chart_histogram(prs, lead_scores, bins=20,
    kicker="Lead-Qualität", headline="Score-Verteilung Mai 2026",
    x_label="Lead-Score")
```

**Scatter-Plot (mit Labels):**
```python
chart_scatter(prs,
    points=[
        (0.78, 0.85, "Galledia"),
        (0.30, 0.60, "Wettbewerb A"),
        (0.55, 0.40, "Wettbewerb B"),
    ],
    x_label="Reichweite", y_label="Tiefe",
    kicker="Positionierung", headline="Galledia oben rechts")
```

**Wann `chart_*` (Sprint 3) vs Sprint 2 Shapes:**
- **Sprint 2** (`waterfall`, `matrix_2x2`, `mekko`, `stacked_bar_pct`, `quadrant_map`) — wenn das Layout-Konzept klar ist und Galledia-CI-Look zwingend
- **Sprint 3** (`chart_*`) — wenn echte Datenmengen visualisiert werden müssen (Histogram, viele Datenpunkte) oder matplotlib-Features (logarithmische Achse, Trend-Annotations) gebraucht werden

Charts werden als PNG (150 DPI) gerendert und in den Slide eingebettet — daher etwas grösserer File-Output (300+ KB pro Chart-Folie).

### Sprint 4 — Strategische Frameworks Teil 1

**SWOT-Analyse (4-Quadranten, MBB-Standard):**
```python
swot(prs,
    strengths=["Marktführer Fachmedien CH", "5 Verlagsobjekte etabliert", "Loyale Kunden"],
    weaknesses=["Print-Anteil noch hoch", "Digital-Ressourcen knapp"],
    opportunities=["Generative KI", "AI Data Center", "Bundle Print + Digital"],
    threats=["Sinkende Print-Werbung", "Plattform-Konkurrenz"],
    kicker="Galledia Fachmedien", headline="SWOT-Analyse 2026")
# Layout: S (TURK) | W (G3)  /  O (RED) | T (G1)
```

**Porter's Five Forces (Zentrum + 4 Aussen-Kräfte mit Pfeilen):**
```python
porter_5f(prs,
    rivalry={"title": "Mittel-hoch", "items": ["3 Mitbewerber", "Preisdruck"]},
    new_entrants={"title": "Niedrig", "items": ["Hohe Hürden"]},
    substitutes={"title": "Hoch", "items": ["Social Media", "Direkt-Marketing"]},
    buyers={"title": "Mittel", "items": ["Konzentrierte Agenturen"]},
    suppliers={"title": "Mittel", "items": ["Papier konzentriert"]},
    kicker="Fachmedien-Branche", headline="Porter's Five Forces 2026")
# Argumente: dict{title, items} ODER tuple(title, list) — beides akzeptiert
```

**Porter's Value Chain (Support oben + Primary unten als Pfeilkette):**
```python
value_chain(prs,
    support_activities=[
        ("Infrastruktur", ["GL", "Controlling", "IT"]),
        ("Personalwesen", ["Recruiting", "Weiterbildung"]),
        ("Technologie", ["CRM", "AI-Pilot"]),
        ("Einkauf", ["Papier", "Lizenzen"]),
    ],
    primary_activities=[
        ("Eingangs-Logistik", ["Content", "Rohstoffe"]),
        ("Operations", ["Redaktion", "Druck"]),
        ("Ausgangs-Logistik", ["Distribution"]),
        ("Marketing & Sales", ["Mediaberatung", "Events"]),
        ("Service", ["Kundenbetreuung"]),
    ],
    kicker="Wertschöpfung", headline="Galledia Value Chain")
# Margin-Label rechts in RED
```

**BCG-Matrix (Portfolio-Analyse, klassisch invertierte X-Achse):**
```python
bcg_matrix(prs,
    products=[
        {"name": "Print A",       "share": 0.85, "growth": 0.20, "revenue": 1200},  # Cash Cow
        {"name": "Digital-Portal","share": 0.30, "growth": 0.80, "revenue":  450},  # Question Mark
        {"name": "Event-Format",  "share": 0.55, "growth": 0.65, "revenue":  380},
        {"name": "Nischen-Print", "share": 0.20, "growth": 0.15, "revenue":  120},  # Dog
    ],
    kicker="Portfolio-Analyse", headline="BCG-Matrix Galledia 2026")
# X-Achse: share=1 (hoch) → LINKS (klassische BCG-Konvention!)
# Bubble-Grösse proportional zu revenue
```

**Wann welches Framework:**
- `swot` → strategischer Überblick, „wo stehen wir gegenüber Markt/Umwelt"
- `porter_5f` → Branchen-Attraktivitäts-Analyse, Wettbewerbsdynamik
- `value_chain` → interne Prozess-Optimierung, „wo erzeugen wir Wert"
- `bcg_matrix` → Portfolio-Entscheidung, „investieren / ernten / verkaufen"

### Sprint 5 — Strategische Frameworks Teil 2

**McKinsey 7S (Shared Values im Zentrum, 6 S's drumherum):**
```python
seven_s(prs, items={
    "shared_values": {"title": "Persönlich"},
    "strategy":      {"title": "Plattform-Bundle"},
    "structure":     {"title": "5 OE / Holding"},
    "systems":       {"title": "Twenty + n8n"},
    "skills":        {"title": "Mediaberatung"},
    "style":         {"title": "Pragmatisch"},
    "staff":         {"title": "ca. 200 MA"},
}, kicker="Organisations-Diagnose", headline="McKinsey 7S Galledia")
# Layout: 6 Kreise hexagonal um zentrales rotes Shared-Values
# Verbindungslinien Zentrum → jedes Outer-S
```

**Ansoff-Matrix (Wachstumsstrategien Produkte × Märkte):**
```python
ansoff_matrix(prs,
    strategies={
        "market_penetration":   ["Bestandskunden-Reaktivierung", "Cross-Sell"],
        "product_development":  ["AI Data Center", "Podcast-Format"],
        "market_development":   ["Mediaagenturen DACH", "Liechtenstein"],
        "diversification":      ["publishr-Lizenzierung", "Event-Plattform"],
    },
    kicker="Wachstums-Strategie", headline="Ansoff-Matrix 2026")
# Risiko-Tags: Niedrig (TURK) → Mittel (G1/G2) → Hoch (RED Diversifikation)
```

**Maturity-Model (Stufentreppe mit Current + Target):**
```python
maturity_model(prs,
    levels=[
        {"name": "Initial",    "description": "Ad-hoc, reaktiv"},
        {"name": "Repeatable", "description": "Erste Standards"},
        {"name": "Defined",    "description": "Dokumentierte Prozesse"},
        {"name": "Managed",    "description": "Messbar gesteuert"},
        {"name": "Optimized",  "description": "Kontinuierl. Verbesserung"},
    ],
    current_level=2, target_level=4,
    kicker="CRM-Reife", headline="Galledia Q1 → Q4 2026")
# Aktuelle Stufe = RED-Marker "▼ HEUTE", Ziel = TURK-Marker "▼ ZIEL"
# Höhe der Stufen wächst linear nach rechts (visuell Treppe)
```

**McKinsey 3-Horizons (überlappende Wachstumskurven):**
```python
three_horizons(prs,
    h1_now      ={"title": "Print + Digital Kern",  "items": ["5 VO", "Bestandskunden"]},
    h2_emerging ={"title": "AI + Plattform",         "items": ["publishr", "AI Data Center"]},
    h3_future   ={"title": "Plattform-Marktplatz",   "items": ["White-Label", "AI-as-a-Service"]},
    kicker="Strategische Roadmap", headline="3-Horizons Galledia")
# H1 (G1/links, jetzt-1J), H2 (TURK/Mitte, 2-4J), H3 (RED/rechts, 4+J)
# Achsen: x=Zeit → / y=Wert ↑
```

**Wann welches Framework:**
- `seven_s` → ganzheitliche Organisations-Diagnose (Strategie+Struktur+Kultur)
- `ansoff_matrix` → Wachstumsoption-Entscheidung mit Risiko-Bewertung
- `maturity_model` → Reifegrad-Assessment (CMMI, ITIL, Digital Maturity)
- `three_horizons` → langfristige Strategie-Roadmap mit Geschäftsmodell-Evolution

### Sprint 6 — Storyline-Enforcement

**Executive Summary — Folie mit nummerierten Key-Insights und Folien-Referenzen:**
```python
exec_summary(prs,
    headline="AI-Hub zahlt sich aus",
    key_points=[
        {"text": "3'900 h/Monat eingespart",   "slide_ref": "S. 7"},
        {"text": "Amortisation in 18 Monaten", "slide_ref": "S. 11"},
        {"text": "5 von 5 Sparten profitieren","slide_ref": "S. 14"},
    ])
# Jede Insight bekommt rote Nummern-Badge, → Referenz dezent rechts
```

**Governing Thought — 1-Satz-Aussage als Banner über Content:**
```python
# Bestehende Folie aus z.B. add_content oder kpi_grid:
s = add_content(prs, "viel", "Markt-Diagnose", "Print sinkt", body, folio="3")

# Governing Thought drauf — zwischen Headline (y≈1.06") und Content (y≈3.30")
add_governing_thought(s, "Die Print-Schwäche wird durch Digital überkompensiert")
# → TURK-Banner mit italic-Text, GW breit, y=2.40"
```

**Section-Tracker — Roadmap-Breadcrumb am oberen Folienrand:**
```python
s = add_content(prs, "viel", "Ausgangslage", "Wo wir heute stehen", body)
add_section_tracker(s,
    sections=["Ausgangslage", "Markt", "Strategie", "Roadmap"],
    current_idx=0)
# Aktive Section: dicker RED-Bar + SB-Label + schwarze Schrift
# Inaktive: dünner G3-Bar + Body-Schrift + G2-Schrift
# Past sections: G2-Bar (zeigt Fortschritt)
```

**SCQA-Struktur — 4-Folien-Pitch-Block (Situation/Complication/Question/Answer):**
```python
slides = scqa_structure(prs,
    situation={
        "headline": "Print dominiert noch",
        "body": "· 65% Umsatz aus Print\n· 5 Verlagsobjekte etabliert\n· Stabile Bestandskunden",
    },
    complication={
        "headline": "Print-Markt schrumpft",
        "body": "· -8% Werbeerlöse 2026\n· -7% Auflagen YoY\n· Cross-Subventionierung nicht nachhaltig",
    },
    question={
        "headline": "Wie reagieren?",
        "body": "· Digital-Investitionen verstärken?\n· Print-Portfolio optimieren?",
    },
    answer={
        "headline": "3-Säulen-Strategie",
        "body": "· Print profitabel halten\n· Digital aggressiv ausbauen\n· AI/Plattform als Horizont 3",
    },
    folio_start="5"  # Folien 5, 6, 7, 8
)
# Erzeugt 4 Folien mit:
# - SCQA-Tracker oben (S→C→Q→A, aktuelle Phase in RED)
# - Code-Badge oben rechts (S/C/Q/A gross, in Phase-spezifischer Farbe)
# - Headline + Body wie add_content
# Returns: list[4 slide objects] für weitere Anpassungen
```

**Wann was nutzen:**
- `exec_summary` → vor dem eigentlichen Inhalt (Folie 2 oder 3), bei GL-/Board-Decks PFLICHT
- `add_governing_thought` → IMMER auf Daten-/Analyse-Folien wo die Schlussfolgerung nicht offensichtlich ist
- `add_section_tracker` → bei Decks ab 12 Folien zur Navigations-Orientierung
- `scqa_structure` → klassische MBB-Storyline-Eröffnung für Pitches und Investor-Decks

### Phase 2 — Erweiterte MBB-Features

**Sprint 8 — Erweiterte Frameworks:**
- `mckinsey_9box(prs, units)` — 9-Box GE-Matrix mit Strategie-Empfehlungen pro Feld
- `bowman_strategy_clock(prs, position, competitors)` — 8 Wettbewerbsstrategien auf Uhr
- `wardley_map(prs, components, edges)` — Wertschöpfung × Reife mit Verbindungen
- `business_model_canvas(prs, segments)` — Osterwalder 9-Felder
- `lewin_change(prs, unfreeze, change, refreeze)` — 3-Phasen Change-Modell
- `kotter_8step(prs, steps, current_step)` — 8-Stufen Change-Process

**Sprint 9 — Advanced Charts:**
- `lever_tornado(prs, levers)` — Sensitivitäts-Analyse (Bars sortiert nach Impact)
- `sankey(prs, flows)` — Flow-Diagramm Quellen → Ziele
- `indexed_comparison(prs, series, base_year)` — Year 1 = 100 Trendvergleich
- `trend_with_confidence(prs, x, mean, lower, upper)` — Prognose mit Band
- `network_diagram(prs, nodes, edges)` — Stakeholder-Map mit Gruppen
- `heatmap_annotated(prs, matrix, row_labels, col_labels, callouts)` — Heat-Map mit Hotspots

**Sprint 10 — Visual Polish:**
- `add_photo_full(slide, image_path, treatment)` — vollflächiges Cover-Photo
  - `treatment`: `none` / `galledia_red_overlay` / `dim` / `dark_vignette`
- `add_photo_card(slide, x, y, w, h, image_path, caption, caption_position)`
  - `caption_position`: `bottom` / `overlay` / `none`
- `add_icon(slide, icon_path, x, y, size_in, color)` — Icon mit PIL-Recoloring
  - `color`: `rot` / `weiss` / `schwarz` / `None`
- `simplify_chart(level)` — matplotlib-Style global setzen
  - `level`: `default` / `light` / `medium` / `heavy` (vor chart_*-Calls)

**Sprint 11 — Quality-Gates:**
- `audit_deck(prs, level)` — Pre-Save-Audit (Generika, Bullet-Count, Quellen, Mix)
  - `level`: `warn` / `strict` / `silent`
- `mece_check(bullet_list)` — Catch-All + Wort-Überlappungs-Detection
- `action_title_check(headline)` — klassifiziert Headline als Action vs Themen-Titel

**Wann was nutzen:**

| Aufgabe | Funktion |
|---|---|
| Portfolio mit Strategie-Empfehlungen | `mckinsey_9box` (feiner als BCG-Matrix) |
| Wettbewerbsstrategie-Visualisierung | `bowman_strategy_clock` |
| Strategische Komponenten-Karte | `wardley_map` |
| Komplettes Geschäftsmodell | `business_model_canvas` |
| Change-Management 3-Phasen | `lewin_change` |
| Change-Process 8 Stufen | `kotter_8step` |
| EBIT-Sensitivität / Hebel-Ranking | `lever_tornado` |
| Mittelfluss (Umsatz → Kosten/Marge) | `sankey` |
| Wachstums-Vergleich indexiert | `indexed_comparison` |
| Prognose mit Unsicherheit | `trend_with_confidence` |
| Stakeholder-/System-Architektur | `network_diagram` |
| Datenmatrix mit Hotspots | `heatmap_annotated` |
| Cover-Folie mit Hintergrund | `add_photo_full` |
| Bild in Layout integriert | `add_photo_card` |
| Icon-Element auf Folie | `add_icon` |
| Tufte-Style Charts (minimal) | `simplify_chart` |
| Quality-Audit vor Save | `audit_deck` (am Ende des Build-Skripts) |
| MECE-Liste prüfen | `mece_check` |
| Headline-Qualität prüfen | `action_title_check` |

**Nicht implementiert (deferred):**
- `add_subsection_tracker`, `add_density_meter`, `enforce_density` (Sprint 12 — geringer Mehrwert)
- `add_content_action_title` (Sprint 13 — Galledia-CI bleibt bei 32-Zeichen-Limit)

Diese Funktionen raisen `NotYetImplementedError` falls aufgerufen.

**Bis ein Sprint abgeschlossen ist:** die entsprechenden Funktionen raisen
`NotYetImplementedError`. Nutze für jetzt die Basis-Layouts und plane Decks
strukturell so, dass die kommenden Layouts sich später einbauen lassen.

---

## Zweck
Ansprechende, abwechslungsreiche Decks im Galledia-CI. CI = Rahmen (Farben, Schrift, Logo,
Layouts). Gestaltung = freie Komposition INNERHALB dieses Rahmens. Niemals ein Einheits-Layout.

---

## Voraussetzungen

| Asset | Pfad | Status |
|---|---|---|
| Template | `assets/Vorlage_6.pptx` | ✅ bereit |
| Volte Regular | `assets/fonts/Volte-Regular.otf` | ✅ bereit |
| Volte Regular Italic | `assets/fonts/Volte-RegularItalic.otf` | ✅ bereit |
| Volte Semibold | `assets/fonts/Volte-Semibold.otf` | ✅ bereit |
| Volte Rounded Regular | `assets/fonts/VolteRounded-Regular.otf` | ✅ bereit |
| Volte Rounded Semibold | `assets/fonts/VolteRounded-Semibold.otf` | ✅ bereit |
| Logo rot (Bildmarke)          | `assets/logo/logo_rot.png`              | ✅ bereit |
| Logo rot + Schriftzug         | `assets/logo/logo_rot_schriftzug.png`   | ✅ bereit |
| Logo weiss (Bildmarke)        | `assets/logo/logo_weiss.png`            | ✅ bereit |
| Logo weiss + Schriftzug       | `assets/logo/logo_weiss_schriftzug.png` | ✅ bereit |
| Logo schwarz (Bildmarke)      | `assets/logo/logo_schwarz.png`          | ✅ bereit |
| Logo schwarz + Schriftzug     | `assets/logo/logo_schwarz_schriftzug.png` | ✅ bereit |
| Helper-Library | `helpers.py` | ✅ bereit |

Tools: `pip install python-pptx Pillow matplotlib --break-system-packages`

---

## Goldene Regel: dichte Folien, klarer Rahmen

Markenwerte: einfach — persönlich — wirkungsvoll. "Einfach" heisst **eine Aussage pro Folie**, nicht **eine Zeile pro Folie**. Jede Folie muss sich beim Vorbeiscrollen wie eine vollständige, abgeschlossene Mini-Story lesen — Headline + Beweis + Detail. Eine Folie mit nur 2 Bullets wirkt unfertig und ist ein Qualitätsfehler.

**Dichte-Soll (verbindlich):**
- Inhaltsfolien (`04_vielText`, `kpi_grid`, `two_column`, `flow_pipeline`, `timeline`, `numbered_steps`): **3–5 Bullets ODER 3–4 KPIs ODER 4–6 Pipeline-Knoten** + Headline + Kapiteltitel + idealerweise Sub-Bullets / Zwischentitel im Body.
- Aussage-Folien (`02_wenigText`): bewusst reduziert — Headline gross + 1–2 Lead-Sätze. NUR für Kernbotschaften, Zwischenfazits, emotionale Anker. Maximal **jede 4. Folie**.
- Titel/Section/Closing: Strukturelemente, zählen nicht zur Dichte-Rechnung.

**Default-Variante = `viel`.** `wenig` ist die Ausnahme, nicht die Regel.

**Treatment nach Inhalt (Markenhandbuch S. 42):**
- informativ / sachlich / Analyse / Status / Roadmap → `04_vielText`, `kpi_grid`, `two_column`, `flow_pipeline`, `timeline`, `numbered_steps` — **das ist der Default**
- emotional / plakativ / Kernbotschaft → `02_wenigText`, Titelfolie, Zwischenfolie — sparsam einsetzen
- Niemals alle Folien gleich. Mindestens 3 verschiedene Layouts pro Deck.

Whitespace ist gestaltet, nicht leer. Leere Bullet-Slots sind kein Whitespace, sondern eine Lücke.

---

## Mindestumfang & Layout-Mix (verbindlich)

**Folienzahl-Untergrenzen:**

| Thema | Min. Folien | Typische Spanne |
|---|---|---|
| Trivial (1 Aussage, internes Statement) | 5 | 5–7 |
| Standard-Update / Statusbericht | 12 | 12–16 |
| Konzept / Strategie / Roadmap | 15 | 15–22 |
| Pitch / Investor / Kunden-Präsi | 16 | 16–25 |

**Pro Hauptkapitel mindestens 2 Inhaltsfolien** (nicht eine Zwischenfolie + eine Inhaltsfolie — das ist zu dünn). Wenn ein Kapitel nur eine Inhaltsfolie hergibt, ist es kein eigenes Kapitel.

**Pflicht-Layout-Mix pro Deck (ab 10 Folien):**
- mindestens **1× `kpi_grid`** (Zahlen kommen IMMER vor, auch in "weichen" Themen — Zeitpunkte, Anzahl Beteiligter, Budget, Termine)
- mindestens **1× `two_column`** (Vergleich Heute/Ziel, Vorher/Nachher, Wir/Sie, Risiko/Chance)
- mindestens **1× `flow_pipeline` ODER `numbered_steps` ODER `timeline`** (Prozess oder Abfolge)
- maximal **60% der Folien dürfen `add_content` sein** — sonst Bleiwüste
- maximal **25% der Folien dürfen `add_content(..., 'wenig', ...)` sein** — sonst Luftnummer

Wenn ein Deck nur `add_content`-Folien hat: **STOP**, Layout-Mix einplanen.

---

## Inhalt zuerst — die wichtigste Regel

Eine Präsentation ist nur so gut wie ihr Inhalt. Eine "mager wirkende" Präsentation hat fast immer eine Inhalts-Ursache, keine Layout-Ursache. Reihenfolge der Inhaltsquellen:

**1. Aktueller Gesprächs-/Projektkontext (höchste Priorität)**
Wenn der Nutzer um eine **Zusammenfassung dieses Projekts / Gesprächs** bittet, ist der Inhalt bereits da — im aktuellen Gespräch und im Projekt-Wissen. Diesen TATSÄCHLICHEN Inhalt zusammenfassen:
- Konkrete Entscheidungen, Schritte, Resultate, Versionen, Zahlen aus dem Gespräch
- Was wurde gebaut, was wurde gelöst, was sind die nächsten Schritte
- Namen, Daten, Versionen, Beträge AKTIV extrahieren — nicht weglassen, weil sie "zu detailliert" wirken
- **Niemals auf generische Aussagen abstrahieren, wenn die Spezifika vor dir liegen.** Eine «Zusammenfassung» die den realen Inhalt durch Wikipedia-Bullets ersetzt, ist ein Totalausfall.

**2. Chat-Historie aktiv durchsuchen (`conversation_search`)**
Bei Galledia-internen Themen, die nicht im aktuellen Gespräch stehen (Jenny, ASMIQ, Digital Twin, n8n, m&k, Press-Release-Pipeline, Archiv): zuerst suchen — echte Zahlen, Architektur, Status, Namen sammeln. **Mehrere Suchen** mit verschiedenen Begriffen, nicht nur eine. Ziel: 10–20 echte Fakten zum Thema.

**3. Projekt-Wissen / Memory durchforsten**
`memory/`-Dateien und CLAUDE.md auf Themen-relevante Einträge prüfen.

**4. Pflichtabfrage**
Nur wenn 1–3 nichts liefern oder das Thema extern ist.

**Inhalts-Tiefe-Test vor dem Bauen:**
Kannst du für jedes geplante Kapitel mindestens 3–4 spezifische Fakten (Zahl, Name, Datum, Betrag, konkrete Entscheidung) nennen? Wenn nein → zurück zu Quelle 1–3, mehr Material sammeln. Niemals mit zu wenig Material in den Build gehen — das produziert exakt die magere Optik, die zu vermeiden ist.

## Pflichtabfrage — IMMER vor dem Bauen

Bevor Code geschrieben wird, diese Fragen stellen — ausser sie sind bereits beantwortet oder aus Gespräch/Historie bekannt:

1. **Datum + Rechtseinheit** (Fusszeile)
2. **Kernbotschaft** — Was soll die Zielgruppe nach der Präsentation denken/tun/entscheiden? (1 Satz)
3. **Zielgruppe + Anlass** — GL-Sitzung, Kundenpräsi, Investor, internes Team? Wie viel wissen sie schon?
4. **Storyline / Kapitelstruktur** — 3–6 Hauptkapitel, je mit Kapiteltitel
5. **Mindestens 8–12 konkrete Fakten/Zahlen/Namen/Zeiträume/Beträge** — das Rohmaterial. Bei <10 Folien reichen 6, bei 15+ Folien sind 12–20 nötig. Beispiele: «3'900h/Monat», «Q3 2026», «Jenny Hostettler», «96 GB VRAM», «CHF 240k», «5 Verlagsobjekte». **Ohne dieses Material wird das Deck generisch.**
6. **Erwarteter Umfang** — wenn der User keine Zahl nennt, vorschlagen nach Tabelle in Abschnitt "Mindestumfang" (Default: 15 Folien für Standard-Themen).

**Wenn der User sagt «mir egal», «nur ein Test», «füll selbst»:** akzeptieren, ABER:
- Chat-Historie + Projekt-Wissen aktiv durchsuchen (`conversation_search`) nach echten Galledia-Fakten zum Thema
- Plausibel-aber-spezifisch erfinden — niemals «Vorteile», «Mehrwert», «Synergien», «Effizienzsteigerung» ohne Zahl dahinter
- Mindest-Folienzahl und Layout-Mix trotzdem einhalten

**Wenn der User explizit «kurzes Deck» / «Übersicht» / «5 Folien» sagt:** seine Vorgabe gilt, Mindestumfang-Tabelle wird übersteuert.


## Qualitätsprinzipien (verbindlich)

Claude leitet den Nutzer zu professionellen, **dicht befüllten** Folien — nicht zu formatierten Aufzählungen, aber auch nicht zu Luft-Folien.

**Pyramid Principle — jede Folie hat eine Hauptaussage:**
```
❌ Titel: «KI-Anwendungen»  (Thema, keine Aussage)
✅ Titel: «KI spart 3'900 Stunden pro Monat»  (Aussage mit Beweis)
```

**Zahlen statt Adjektive — generische Begriffe sind verboten:**
```
❌ «Signifikante Effizienzsteigerung»
❌ «Vorteile», «Mehrwert», «Synergien», «Best Practices», «Optimierung», «Transformation» — als Standalone-Bullet
✅ «39% weniger manuelle Arbeit — 3'900h/Monat»
✅ «Jenny Hostettler übernimmt Q3 2026 die ASMIQ-Pipeline»
```
Wenn ein Bullet ohne Zahl/Name/Datum/Betrag auskommt, gehört er meistens nicht aufs Slide. Ausnahme: Storyline-Anker und Zwischentitel.

**Inhaltsdichte pro Folie (verbindlich):**
```
❌ Folie mit nur Headline + 2 Bullets  → wirkt leer, Qualitätsfehler
✅ Headline + Kapiteltitel + 3–5 Bullets + ggf. Sub-Bullets / Zwischentitel
✅ Headline + 3–4 KPI-Callouts mit Zahl + Label
✅ Headline + 2 Spalten à 3–4 Bullets
✅ Headline + 4–6 Pipeline-Knoten
```

**Sub-Bullets aktiv nutzen** — verdoppeln die Informationsdichte ohne neue Folie:
```
• Stammdaten ASMIQ
•• Company, Person mit Custom-Feldern (Suchname, MwSt-Nr)
•• 5 Verlagsobjekte mit Rabattstaffeln und Beraterkommission
• Opportunity-Pipeline
•• Angebotsnummern Format 2026-00001
•• Bundle-Pakete mit Listenpreis, Rabatt, Endpreis
```

**Visuelle Beweise statt Bullet-Listen — bei jeder passenden Gelegenheit:**
```
→ Zahlen → kpi_grid()
→ Vergleich → two_column()
→ Prozess → flow_pipeline() / numbered_steps()
→ Zeitverlauf → timeline()
→ Statement → image_bleed() oder add_content('wenig', ...)
```

**Eine Aussage pro Folie — aber mit Beweis:**
Die Aussage steht in der Headline. Die 3–5 Bullets / KPIs / Schritte sind der BEWEIS dafür, nicht weitere Aussagen. So entstehen dichte Folien mit klarer Hierarchie statt 6 gleichrangige Bullets.

**Roter Faden — typische Struktur für 15-Folien-Deck:**
```
Folie 1:  Titel (Thema + Kernbotschaft im Untertitel)
Folie 2:  Agenda (4–6 Kapitel)
Folie 3:  Zwischenfolie Kapitel 1 — Ausgangslage
Folie 4:  Ausgangslage / Status quo (add_content viel, 4 Bullets mit Zahlen)
Folie 5:  KPI-Grid mit 3–4 Schlüsselzahlen
Folie 6:  Zwischenfolie Kapitel 2 — Problem
Folie 7:  Problem-Analyse (two_column Heute/Lücke)
Folie 8:  Wirkung des Problems (add_content viel mit Sub-Bullets)
Folie 9:  Zwischenfolie Kapitel 3 — Lösung
Folie 10: Lösungskonzept (flow_pipeline 4–5 Knoten)
Folie 11: Lösungsdetail (add_content viel)
Folie 12: Zwischenfolie Kapitel 4 — Umsetzung
Folie 13: Roadmap (timeline mit 4–5 Phasen)
Folie 14: Nächste Schritte (numbered_steps mit 3–4 Schritten)
Folie 15: Schlussfolie
```

## Setup

```bash
pip install python-pptx Pillow matplotlib --break-system-packages
```

```python
# Assets von GitHub laden (Verzeichnisstruktur für helpers.py erhalten)
import os, sys, urllib.request
_DIR = "/tmp/galledia_praesentation"
os.makedirs(f"{_DIR}/assets/logo", exist_ok=True)
sys.path.insert(0, _DIR)
_BASE = "https://galledia-ag.github.io/galledia-office-ci/skills/galledia-praesentation"
_FILES = [
    "helpers.py",
    "assets/Vorlage_6.pptx",
    "assets/logo/logo_rot.png",
    "assets/logo/logo_rot_schriftzug.png",
    "assets/logo/logo_weiss.png",
    "assets/logo/logo_weiss_schriftzug.png",
    "assets/logo/logo_schwarz.png",
    "assets/logo/logo_schwarz_schriftzug.png",
]
for _name in _FILES:
    _dest = f"{_DIR}/{_name}"
    if not os.path.exists(_dest):
        try:
            urllib.request.urlretrieve(f"{_BASE}/{_name}", _dest)
            print(f"✓ {_name} ({os.path.getsize(_dest):,} bytes)")
        except Exception as e:
            # HARD-STOP: Niemals from-scratch improvisieren ohne CI-Assets.
            raise RuntimeError(
                f"Asset-Download fehlgeschlagen ({_name}): {e}. "
                f"Plugin/Repo nicht erreichbar — Skill abbrechen und Asset-Fehler "
                f"an User melden. NIEMALS mit python-pptx ohne Vorlage_6/Volte/"
                f"Logos eine 'Galledia'-Präsentation bauen (CI-Verstoss garantiert)."
            )

from helpers import (build_presentation, add_title, add_section, add_agenda,
                     add_content, add_closing, add_discussion,
                     kpi_grid, two_column, flow_pipeline, numbered_steps, timeline)
```


## Pflichtregeln — Häufige Fehler

**add_agenda():**
```python
# ❌ FALSCH — 2 Zeilen pro Punkt, Teaser verboten
add_agenda(prs, ["KI-Anwendungen
Status und Hintergrund", "Roadmap
Massnahmen"])
# ✅ RICHTIG — 1 Zeile pro Punkt, kein Zusatz
add_agenda(prs, ["KI-Anwendungen", "Roadmap", "Nächste Schritte"], folio="2")
```

**add_content() — Kapiteltitel:**
```python
# ❌ FALSCH — Zahlen als Kapiteltitel
add_content(prs, "viel", "01", "Headline", "Text", folio="3")
# ✅ RICHTIG — beschreibendes Wort
add_content(prs, "viel", "KI-Trends", "Headline max. 32 Zeichen", "Text", folio="3")
```

**kpi_grid() / two_column() / flow_pipeline() — kicker + headline PFLICHT:**
```python
# ❌ FALSCH — kicker und headline fehlen → leerer Seitenkopf
two_column(prs, "Heute", items_l, "Ziel", items_r, folio="4")
# ✅ RICHTIG
two_column(prs, "Heute", items_l, "Ziel 2026", items_r,
           kicker="Transformation", headline="Heute vs. Ziel 2026",
           col2_red=True, folio="4")

# ❌ FALSCH
kpi_grid(prs, [("96 GB","VRAM")], folio="3")
# ✅ RICHTIG
kpi_grid(prs, [("96 GB","VRAM"), ("256 GB","RAM"), ("6 TB","NVMe")],
         kicker="Hardware", headline="R&D AI-Workhorse", folio="3")
```

**Längen-Limits (HART, vom Code erzwungen):**

| Parameter | Max-Zeichen | Wo |
|---|---|---|
| `headline` (Argument zu `add_content` / `kpi_grid` / `two_column` etc.) | **32** | bei 72pt auf 1 Zeile |
| `kicker` / `kapitel` (Kapiteltitel-Argument, immer als 2. Positional) | **35** | bei 30pt auf 1 Zeile |
| `title` von `add_title()` | **40** für 1-Zeilen-Optik (≤40 = 1 Zeile, >40 = 2 Zeilen) |
| Agenda-Items | **~55** je Punkt, 1 Zeile |

⚠️ **Wichtig für LLM-Generierung:** Plane Headlines von Anfang an **knackig und unter 32 Zeichen**. Der Code raised ValueError bei Überschreitung — du musst sonst die ganze Präsentation neu generieren.

```python
# ❌ FALSCH — 37 Zeichen, ValueError
kpi_grid(prs, kpis, headline="Start ins neue Geschäftsjahr gelungen", ...)
# ✅ RICHTIG — 22 Zeichen
kpi_grid(prs, kpis, headline="Gelungener Jahresstart", ...)

# ❌ FALSCH — 34 Zeichen
two_column(prs, ..., headline="Turnaround Print — vier Baustellen", ...)
# ✅ RICHTIG — 27 Zeichen
two_column(prs, ..., headline="Print dreht — vier Baustellen", ...)
```

**Tipp für knackige Headlines:** Aussage statt Beschreibung, Verben statt Substantive, weglassen statt erklären.

**body_text-Format für `add_content()` — strikte Konvention:**

Der Renderer parst das body_text-String zeilenweise und unterscheidet drei Zeilen-Typen:

| Zeile beginnt mit | Rendering |
|---|---|
| `• `, `- `, `* `, oder `· ` | **Bullet Ebene 1** — Volte 19pt Regular, schwarz, mit echtem PowerPoint-Bullet `•` |
| `•• `, `-- `, oder `·· ` (oder Einrückung + Bullet) | **Bullet Ebene 2** — Volte 17pt Regular, grau, mit `–` und Einrückung |
| Beliebiger Text **ohne** Bullet-Marker | **Zwischentitel** — Volte Semibold 22pt, schwarz, Abstand davor |
| Leerzeile | kleiner Abstand |

`**bold**`-Markdown wird automatisch gestrippt (Bold rendern wir via Schrift, nicht via Marker). Du brauchst Sterne also nicht — und falls sie reinrutschen, schaden sie nicht.

```python
# ✅ RICHTIG — alle drei Bullet-Marker funktionieren gleich:
body = """Stammdaten
• Company, Person — mit Custom-Feldern (Suchname, MwSt-Nr)
• Verlagsobjekt — Stammdaten, Rabattstaffeln, Beraterkommission

Verkauf
• Opportunity mit Angebotsnummer (Format 2026-00001)
• Angebotspaket — Bundle mit Listenpreis, Rabatt, Endpreis
"""
# «Stammdaten» und «Verkauf» werden als Zwischentitel (Semibold 22pt) gerendert,
# die `• `-Zeilen als echte Bullets (Regular 19pt).
```

```python
# ❌ FALSCH — `**bold**` ist unnötig (wird gestrippt), aber ASCII-Bullets fehlen:
body = "**Stammdaten**\nCompany, Person — ohne Marker bleibt's Semibold-Zwischentitel"
# → die zweite Zeile wird fälschlich als weiterer Zwischentitel gerendert.
```


---

## Workflow (4 Schritte)

**Pflicht vor dem ersten `build_presentation()`-Aufruf:**
Wenn Datum und/oder Rechtseinheit im Gespräch nicht bekannt sind → User fragen:
- «Für welches Datum soll die Fusszeile gesetzt werden? (Beispiel: 29. Mai 2026)»
- «Für welche Rechtseinheit? (Beispiel: Galledia Fachmedien AG)»
Erst nach Antwort fortfahren.

```python
from helpers import build_presentation, add_title, add_section, add_content
from helpers import kpi_grid, two_column, flow_pipeline, numbered_steps, add_closing

prs = build_presentation(          # Datum + Rechtseinheit IMMER setzen
    datum="29. Mai 2026",          # ← vom User bestätigt
    rechtseinheit="Galledia Fachmedien AG"  # ← vom User bestätigt
)
add_title(prs, "Titel", "Untertitel | Datum")
add_section(prs, "01", "Ausgangslage")
add_content(prs, "viel", "Kapiteltitel", "Headline", "Textkörper...", folio="3/12")
kpi_grid(prs, [("96 GB","VRAM"), ("256 GB","RAM"), ("6 TB","NVMe")])
two_column(prs, "Heute", bullets_l, "Ziel 2026", bullets_r, col2_red=True)
flow_pipeline(prs, ["Netzlaufwerk","Parsing","Vektor-DB","Antwort"])
numbered_steps(prs, [("Hardware","..."), ("Deployment","..."), ("Indizierung","...")])
add_closing(prs)                    # Schlussfolie (rotes Layout, weisses Logo)
prs.save("output.pptx")
```

**Farbrhythmus:** rote Zwischenfolien als Kapitel-Anker → helle Inhaltsfolien dazwischen.
Nie zwei Zwischenfolien hintereinander, nie alle Folien gleich.

---

## Layout-Entscheidungstabelle (Vorlage_6)

| Inhaltstyp | Layout-Name | Platzhalter |
|---|---|---|
| Deck-Titel | `Titelfolie` | idx=0 Titel, idx=10 Untertitel |
| Kapitel-Anker / Abschnitt | `2_Zwischenfolie rot` | keine — Textbox manuell |
| Agenda wenig (≤5 Punkte) | `01_Agenda 5` | idx=0 Label, idx=11 leer, idx=13 Punkte, idx=14 Folio |
| Agenda viel (6–12 Punkte) | `01_Agenda 22` | idx=0 Label, idx=11 leer, idx=13 Punkte, idx=14 Folio |
| Grosse Aussage / Kernbotschaft | `02_wenigText` | idx=0 Kapiteltitel, idx=11 Headline 72pt, idx=13 Lead, idx=14 Folio, idx=15 Quelle |
| Detail / Text / Struktur | `04_vielText` | idx=0 Kapiteltitel, idx=11 Headline, idx=13 Textkörper, idx=14 Folio, idx=15 Quelle |
| KPI, Zweispalter, Pipeline, Timeline | `Leer` | idx=0/11/13 leer lassen, idx=14 Folio — Shapes via helpers.py |
| Diskussion / Interaktion | `Abschlussfolie` | keine — zeigt Piktogramm + «Diskussion» |
| Piktogramm-Referenz | `Piktogramme` | idx=0 Titel |
| Schluss-/Abschlussfolie | `Schlussfolie` | keine — zeigt weisses Logo auf Rot |

**Hinweis `Leer`:** Im Bearbeitungsmodus zeigt das Layout leere Platzhalter-Rahmen
(Bekanntes Problem aus Vorlage_6). In Render + Präsentation unsichtbar. Kein Blocker.

---

## CI-Farb-Tokens (aus Template-Theme — verbindlich)

```python
# helpers.py exportiert diese als Konstanten
RED   = RGBColor(0xE6, 0x1C, 0x52)   # Galledia-Rot  — Primärfarbe
BLACK = RGBColor(0x00, 0x00, 0x00)   # Galledia-Schwarz
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
G1    = RGBColor(0x40, 0x40, 0x40)   # Grau 1 (64/64/64)
G2    = RGBColor(0x66, 0x66, 0x66)   # Grau 2 (102/102/102)
G3    = RGBColor(0xA6, 0xA6, 0xA6)   # Grau 3
G4    = RGBColor(0xD9, 0xD9, 0xD9)   # Grau 4
GL    = RGBColor(0xF2, 0xF2, 0xF5)   # Hintergrundgrau (Karten)
TURK  = RGBColor(0x22, 0xAA, 0x9F)   # Türkis — NUR Info-/Quote-Panels
```

**Gewichtung:** dominant Rot/Schwarz/Grau. Türkis, Purple, Blau, Bronze NUR für Infografiken
und Akzente — nie als Hauptfarbe.

---

## CI-Schrift-Tokens

| Verwendung | Font-Name |
|---|---|
| Plakative Titel / Headlines | `Volte Rounded Semibold` |
| Fliesstext / Lead / Labels | `Volte` |
| Zwischentitel / Hervorhebung | `Volte Semibold` |
| Italic (Hervorhebung im Text) | `Volte` + italic=True |

Ausrichtung: **linksbündig** auf Inhaltsfolien. **Zentriert** nur auf Titelfolie/Cover.
Body-Text grundsätzlich in Schwarz.

---

## CI-Mikroregeln (verbindlich)

- **Logo:** nur Rot, Weiss oder Schwarz. Nie mit Box/Umform auf Bildern. Nie verzerren.
  Schutzzone wahren. Unter 8 mm Breite → nur Bildmarke (G), ohne Schriftzug.
  **Farbwahl:** Rot-Logo auf weissem/hellem Hintergrund · Weiss-Logo auf rotem/dunklem Hintergrund · Schwarz-Logo auf hellem Hintergrund (alternative zu Rot).
  **Skill-Funktion:** `add_logo(slide, variant='rot')` bzw. `variant='weiss'` — platziert Bildmarke rechts unten.
- **Bullets:** Aufzählungszeichen `•` (Punkt). Zitate in Guillemets `« »`. Separator `|`.
- **Formen:** Kästen und Linien immer mit abgerundeten Ecken (Motiv konsistent).
- **Bildwelt:** authentische Menschen «wie du und ich», warmes weiches Licht, erdige Töne,
  partielle Highlights in Galledia-Rot. Keine offensichtlichen Stock-Models.
- **Piktogramme:** plakativ (gross/dominant) oder informativ (klein). In Schwarz/Weiss/Rot
  oder Akzentfarbe. Piktogramme NICHT mit dem Galledia-Alphabet mischen.
- **Fusszeile (idx=14):** Format `«n / total»` z.B. `«3 / 12»`. Auf Titelfolie und
  Schlussfolie weglassen.
- **Quellenangabe (idx=15):** nur setzen wenn Quellen vorhanden. Format:
  `Quelle: [Quelle 1] / [Quelle 2, Jahr]`.

---

## Anti-Bleiwüste UND Anti-Luftnummer (Pflichtprüfung pro Folie)

Beide Extreme sind Fehler. Pro Folie prüfen:

**Gegen Bleiwüste:**
- Jede Folie hat ≥1 visuelles Element (Bild, Piktogramm, grosse Zahl, Diagramm, Farbfläche, Pipeline-Shapes).
- Max. **8 Zeilen** Body-Text (inkl. Sub-Bullets) — bei mehr splitten.
- Max. **5 Top-Level-Bullets** — sonst splitten oder in `two_column`.
- Vergleich/Prozess/Timeline als Shapes statt als Bullet-Liste.

**Gegen Luftnummer:**
- Min. **3 Bullets ODER 3 KPIs ODER 4 Pipeline-Knoten** auf Inhaltsfolien — bei weniger entweder dichter füllen oder mit Nachbarfolie zusammenlegen.
- Min. **eine konkrete Zahl/Name/Datum** auf jeder Sach-Inhaltsfolie (Ausnahmen: Section, Closing, bewusste Aussage-Folie mit `02_wenigText`).
- KEIN Bullet, der nur aus generischen Begriffen besteht (`Vorteile`, `Effizienz`, `Mehrwert`, `Synergien`, `Optimierung`, `Best Practices`, `Skalierung` ohne Kontext).
- KEIN Kapitel mit nur 1 Inhaltsfolie nach der Zwischenfolie — dann ist es kein Kapitel.

**Rhythmus:**
- Roter Farbrhythmus: mindestens jede 3.–4. Folie ein Rot-Anker (Zwischenfolie oder rote Akzentfolie).
- Layout-Rhythmus: in 10 aufeinanderfolgenden Inhaltsfolien max. 6× `add_content` — Rest aus `kpi_grid` / `two_column` / `flow_pipeline` / `timeline` / `numbered_steps` / `image_bleed`.
- KEINE Akzentlinien unter Titeln. KEINE dekorativen Vollbalken. Vollflächige rote Cover/Zwischenfolien sind dagegen CI-Signatur und ausdrücklich erwünscht.

---

## Quellenangabe und Fusszeile setzen

```python
# In native Layouts (04_vielText, 02_wenigText):
ph[14].text = "«5 / 12»"                                        # Folienzahl
ph[15].text = "Quelle: Gartner 2025 / Schätzung Galledia F&E"  # optional

# Auf Leer (via helpers.add_footer):
add_footer(slide, folio="5 / 12", source="Gartner 2025")
```

---

## QA (verbindlich vor Abgabe)

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg && pdftoppm -jpeg -r 120 output.pdf slide
```

Subagent-Inspektion: Textüberlauf, Kontrast, Überlappung, Schriftsubstitution, Platzhalter-Reste.

```bash
extract-text output.pptx | grep -iE "xxx|lorem|Mastertitel|Kapiteltitel 30pt|Themenpunkt"
```

---

## Erweiterung: Word-Vorlagen (Phase 2)

Brief, Kurzbrief und Dokument folgen in einer zweiten Skill-Sektion (`## Word-Vorlagen`).
Selbe CI-Tokens, selbe Fonts, python-docx als Builder. Sobald Word-Templates geliefert werden.
