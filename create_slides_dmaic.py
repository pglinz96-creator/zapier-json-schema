from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── Light Design Palette ────────────────────────────────────────
BG          = RGBColor(0xF7, 0xF9, 0xFC)   # near-white background
DARK        = RGBColor(0x1A, 0x20, 0x2C)   # primary text
MID         = RGBColor(0x4A, 0x55, 0x68)   # secondary text
LIGHT_BOX   = RGBColor(0xED, 0xF2, 0xF7)   # light grey box
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE      = RGBColor(0xE8, 0x5D, 0x04)   # Schellenberg accent

# DMAIC phase colours
D_COL = RGBColor(0x31, 0x82, 0xCE)  # blue
M_COL = RGBColor(0x80, 0x5A, 0xD5)  # purple
A_COL = RGBColor(0xDD, 0x6B, 0x20)  # orange-red
I_COL = RGBColor(0x38, 0xA1, 0x69)  # green
C_COL = RGBColor(0x00, 0x96, 0xB5)  # teal

PHASE_COLORS = {"D": D_COL, "M": M_COL, "A": A_COL, "I": I_COL, "C": C_COL}
PHASE_BG = {
    "D": RGBColor(0xEB, 0xF4, 0xFF),
    "M": RGBColor(0xF3, 0xEE, 0xFD),
    "A": RGBColor(0xFF, 0xF0, 0xE6),
    "I": RGBColor(0xEA, 0xF7, 0xEE),
    "C": RGBColor(0xE0, 0xF7, 0xFA),
}

def set_bg(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, text, l, t, w, h,
                font_size=14, bold=False, color=DARK,
                align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(
        Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_rect(slide, l, t, w, h, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1.2)
    else:
        shape.line.fill.background()
    return shape

def add_multiline(slide, lines, l, t, w, h,
                  font_size=13, color=DARK, bold_first=False):
    txBox = slide.shapes.add_textbox(
        Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        if bold_first and line == lines[0]:
            run.font.bold = True

def phase_header(slide, phase_letter, phase_name, subtitle):
    """Top bar with phase colour + title."""
    col = PHASE_COLORS[phase_letter]
    add_rect(slide, 0, 0, 13.33, 1.15, col)
    add_textbox(slide, phase_letter, 0.25, 0.08, 0.9, 0.9,
                font_size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, phase_name, 1.3, 0.1, 8, 0.55,
                font_size=26, bold=True, color=WHITE)
    add_textbox(slide, subtitle, 1.3, 0.68, 10, 0.38,
                font_size=13, color=RGBColor(0xDD, 0xEE, 0xFF))
    # Company tag
    add_textbox(slide, "Schellenberg Druck AG", 9.5, 0.1, 3.6, 0.38,
                font_size=11, color=WHITE, align=PP_ALIGN.RIGHT)

def card(slide, x, y, w, h, title, items, phase="D",
         item_color=DARK, font_size=12):
    col = PHASE_COLORS[phase]
    bg  = PHASE_BG[phase]
    add_rect(slide, x, y, w, h, bg, line_color=col)
    add_rect(slide, x, y, w, 0.42, col)
    add_textbox(slide, title, x+0.15, y+0.04, w-0.3, 0.35,
                font_size=12, bold=True, color=WHITE)
    add_multiline(slide, items, x+0.15, y+0.5, w-0.3,
                  h-0.6, font_size=font_size, color=item_color)

# ══════════════════════════════════════════════════════════════════
# SLIDE 1 – TITLE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)

# Left accent bar
add_rect(slide, 0, 0, 0.5, 7.5, ORANGE)

# Top light strip
add_rect(slide, 0.5, 0, 12.83, 0.08, LIGHT_BOX)

# DMAIC phase circles (decorative)
for i, (letter, col) in enumerate([
    ("D", D_COL), ("M", M_COL), ("A", A_COL), ("I", I_COL), ("C", C_COL)
]):
    cx = 1.2 + i * 2.2
    add_rect(slide, cx, 5.5, 1.5, 1.5, col)
    add_textbox(slide, letter, cx, 5.55, 1.5, 1.4,
                font_size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Titles
add_textbox(slide, "KI-gestützte Auftragsannahme", 1.2, 1.2, 11, 1.1,
            font_size=46, bold=True, color=DARK)
add_textbox(slide, "Prozessoptimierung nach DMAIC", 1.2, 2.45, 11, 0.65,
            font_size=22, color=ORANGE, bold=True)
add_textbox(slide, "Schellenberg Druck AG · Schützenhausstrasse 5 · 8330 Pfäffikon",
            1.2, 3.2, 11, 0.5, font_size=15, color=MID)
add_textbox(slide, "n8n · Claude AI · Gmail · Google Docs · Multi-Channel",
            1.2, 3.85, 11, 0.45, font_size=14, color=MID)
add_textbox(slide, "Juni 2026", 1.2, 6.9, 11, 0.4,
            font_size=12, color=MID, align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════════
# SLIDE 2 – DMAIC ÜBERSICHT
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_rect(slide, 0, 0, 13.33, 0.08, ORANGE)

add_textbox(slide, "DMAIC Projektübersicht", 0.6, 0.2, 12, 0.65,
            font_size=28, bold=True, color=DARK)
add_textbox(slide, "KI-Auftragsannahme Schellenberg Druck AG",
            0.6, 0.88, 12, 0.38, font_size=14, color=MID)

phases = [
    ("D", "DEFINE",   D_COL, PHASE_BG["D"], [
        "Problem: Manuelle Auftragsannahme",
        "Ziel: Vollautomatisierung",
        "Scope: Alle Eingangskanäle",
        "KPIs definiert",
    ]),
    ("M", "MEASURE",  M_COL, PHASE_BG["M"], [
        "Ist-Zustand: 30–60 min/Auftrag",
        "Fehlerquellen gemessen",
        "Kanalvolumen analysiert",
        "Baseline festgelegt",
    ]),
    ("A", "ANALYZE",  A_COL, PHASE_BG["A"], [
        "Ursachen: fehlende Standards",
        "Manuelle Kalkulation fehleranfällig",
        "Thread-Kontext geht verloren",
        "Mehrprodukt-Aufträge komplex",
    ]),
    ("I", "IMPROVE",  I_COL, PHASE_BG["I"], [
        "n8n + Claude AI Workflow",
        "6 Eingangskanäle normalisiert",
        "Deterministisches Pricing",
        "HitL mit Wait-Node",
    ]),
    ("C", "CONTROL",  C_COL, PHASE_BG["C"], [
        "3-Ebenen Fehlerbehandlung",
        "Edgar prüft & gibt frei",
        "Audit-Trail via Google Docs",
        "Monitoring & Alerts",
    ]),
]

bw = 2.3
for i, (letter, name, col, bg, items) in enumerate(phases):
    x = 0.5 + i * (bw + 0.12)
    add_rect(slide, x, 1.55, bw, 5.45, bg, line_color=col)
    add_rect(slide, x, 1.55, bw, 0.75, col)
    add_textbox(slide, letter, x+0.1, 1.6, 0.7, 0.65,
                font_size=26, bold=True, color=WHITE)
    add_textbox(slide, name, x+0.75, 1.68, bw-0.85, 0.5,
                font_size=14, bold=True, color=WHITE)
    add_multiline(slide, items, x+0.15, 2.45, bw-0.3,
                  4.3, font_size=12, color=MID)
    if i < 4:
        add_textbox(slide, "→", x+bw, 3.65, 0.15, 0.5,
                    font_size=16, bold=True, color=col, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 3 – DEFINE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "D", "DEFINE – Problemdefinition", "Was wollen wir lösen und warum?")

# Problem box
card(slide, 0.5, 1.35, 5.9, 2.8, "Ausgangsproblem", [
    "Druckaufträge per E-Mail, Telefon, Web manuell erfasst",
    "30–60 Minuten Bearbeitungszeit pro Auftrag",
    "Preise von Hand kalkuliert – hohe Fehlerquote",
    "Angebote ohne einheitliche Vorlage erstellt",
    "Keine Skalierbarkeit bei steigendem Auftragsvolumen",
], phase="D", font_size=12)

# Goal box
card(slide, 6.7, 1.35, 5.9, 2.8, "Projektziel", [
    "Vollautomatische Auftragsannahme über alle Kanäle",
    "KI extrahiert Daten, prüft Vollständigkeit",
    "Deterministische Preiskalkulation (keine Spielräume)",
    "AI-generiertes, rechtssicheres Angebot (CH)",
    "Mensch prüft nur noch → gibt frei oder lehnt ab",
], phase="D", font_size=12)

# Scope + KPIs
card(slide, 0.5, 4.35, 5.9, 2.85, "Scope & Stakeholder", [
    "Scope: Alle eingehenden Druckaufträge",
    "Kanäle: Mail, Web, WhatsApp, Telegram,",
    "         Telefon, Kontaktformular",
    "Stakeholder: Edgar Glinz (Entscheider)",
    "             Kunden der Schellenberg AG",
    "Ausgeschlossen: Produktion, Rechnungsstellung (Phase 3)",
], phase="D", font_size=12)

card(slide, 6.7, 4.35, 5.9, 2.85, "Erfolgskennzahlen (KPIs)", [
    "Bearbeitungszeit: < 5 Min. (von 30–60 Min.)",
    "Kalkulationsfehlerrate: 0% (deterministische Logik)",
    "Vollständige Aufträge ohne Rückfrage: > 70%",
    "Angebotserstellungszeit: < 2 Min.",
    "HitL-Freigabequote: Ziel > 90% genehmigt",
], phase="D", font_size=12)

# ══════════════════════════════════════════════════════════════════
# SLIDE 4 – MEASURE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "M", "MEASURE – Ist-Zustand", "Wo stehen wir heute? Was kostet der aktuelle Prozess?")

card(slide, 0.5, 1.35, 3.9, 5.85, "Ist-Prozess (manuell)", [
    "1.  E-Mail / Anruf eingehend",
    "2.  Sachbearbeiter liest & interpretiert",
    "3.  Fehlende Angaben per Mail anfragen",
    "4.  Warten auf Kundenantwort",
    "5.  Preise in Excel kalkulieren",
    "6.  Word-Vorlage ausfüllen",
    "7.  PDF erstellen & versenden",
    "8.  Auf Freigabe/Rückfragen warten",
    "",
    "⏱ Ø Zeitaufwand: 30–60 Min./Auftrag",
    "⚠  Fehlerquellen: Schritt 2, 3, 5, 6",
], phase="M", font_size=12)

card(slide, 4.65, 1.35, 3.9, 2.7, "Kanäle & Volumen", [
    "E-Mail         ~60% der Anfragen",
    "Telefon        ~25% der Anfragen",
    "Web-Formular   ~10% der Anfragen",
    "Andere         ~5% der Anfragen",
    "",
    "Ziel: Alle Kanäle automatisiert",
], phase="M", font_size=12)

card(slide, 4.65, 4.25, 3.9, 2.95, "Fehlerquellen-Analyse", [
    "Falsche Grammatur → Preisfehler",
    "Thread-Kontext vergessen",
    "Format-Tippfehler (A4 vs. DIN lang)",
    "Mehrprodukt-Aufträge übersehen",
    "Keine Fristenprüfung",
], phase="M", font_size=12)

card(slide, 8.8, 1.35, 4.3, 5.85, "Baseline Kennzahlen", [
    "Bearbeitungszeit heute:",
    "  → Einfacher Auftrag:    ~30 Min.",
    "  → Komplexer Auftrag:    ~60 Min.",
    "  → Mit Rückfragen:      ~2–3 Tage",
    "",
    "Fehlerrate Kalkulation:   ~15–20%",
    "(geschätzt, keine Messung)",
    "",
    "Angebote pro Woche: variabel",
    "",
    "Ziel nach Verbesserung:",
    "  → Bearbeitungszeit: < 5 Min.",
    "  → Fehlerrate Kalk.: 0%",
    "  → Rückfragen: automatisch",
], phase="M", font_size=12)

# ══════════════════════════════════════════════════════════════════
# SLIDE 5 – ANALYZE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "A", "ANALYZE – Ursachenanalyse", "Warum ist der Prozess fehleranfällig und langsam?")

root_causes = [
    ("Keine Standardisierung", A_COL, [
        "Kunden liefern Infos in beliebigem Format",
        "Kein einheitliches Bestellformular",
        "Jeder Sachbearbeiter arbeitet anders",
        "Keine Vollständigkeits-Checkliste",
    ]),
    ("Manuelle Kalkulation", A_COL, [
        "Preistabellen in Excel – fehleranfällig",
        "Expresszuschläge oft vergessen",
        "Marge nicht systematisch geprüft",
        "Mehrprodukt-Summen falsch addiert",
    ]),
    ("Informationsverlust", A_COL, [
        "E-Mail-Thread wird nicht vollständig gelesen",
        "Frühere Angaben gehen bei Rückfragen verloren",
        "Kein zentrales Auftragssystem",
        "Anhänge werden übersehen",
    ]),
]

for i, (title, col, items) in enumerate(root_causes):
    x = 0.5 + i * 4.2
    card(slide, x, 1.35, 4.0, 3.6, title, items, phase="A", font_size=12)

# Solution mapping
add_rect(slide, 0.5, 5.15, 12.3, 2.1, PHASE_BG["A"], line_color=A_COL)
add_rect(slide, 0.5, 5.15, 12.3, 0.42, A_COL)
add_textbox(slide, "Ableitung: Lösungsansatz für IMPROVE-Phase", 0.65, 5.2,
            11.5, 0.32, font_size=12, bold=True, color=WHITE)

solutions = [
    ("Standardisierung",   "KI-Extraktion + JSON-Schema erzwingt Vollständigkeit"),
    ("Kalkulation",        "Fixe Preistabellen im Prompt – kein Ermessensspielraum"),
    ("Thread-Kontext",     "AI liest gesamten E-Mail-Verlauf, merkt sich Auftragsnummer"),
    ("Multi-Kanal",        "6 Kanäle werden normalisiert auf identisches JSON-Format"),
]
for i, (prob, sol) in enumerate(solutions):
    col_x = 0.7 + i * 3.05
    add_textbox(slide, prob, col_x, 5.65, 2.9, 0.35,
                font_size=11, bold=True, color=A_COL)
    add_textbox(slide, sol, col_x, 6.05, 2.9, 0.9,
                font_size=11, color=MID)

# ══════════════════════════════════════════════════════════════════
# SLIDE 6 – IMPROVE (Lösung Übersicht)
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "I", "IMPROVE – Lösung: n8n KI-Workflow", "Von der Idee zur automatisierten Auftragsannahme")

# Main workflow flow
steps_improve = [
    ("📥", "Multi-Channel\nInput",      "6 Kanäle\nnormalisiert"),
    ("🤖", "Vollständig-\nkeitsprüfung","Claude AI\n+ Schema"),
    ("💰", "Preis-\nkalkulation",       "Deterministisch\nFixpreise"),
    ("📄", "Offerte\ngenerieren",       "Google Docs\nCH-konform"),
    ("👁", "HitL\nFreigabe",            "Wait-Node\nEdgar entsch."),
    ("✉️", "Versand\nKunde",            "Angebot oder\nNachfrage"),
]
bw3, bh3 = 1.85, 2.5
for i, (icon, title, desc) in enumerate(steps_improve):
    x = 0.6 + i * (bw3 + 0.17)
    add_rect(slide, x, 1.35, bw3, bh3, PHASE_BG["I"], line_color=I_COL)
    add_rect(slide, x, 1.35, bw3, 0.4, I_COL)
    add_textbox(slide, icon, x, 1.42, bw3, 0.32, font_size=16,
                align=PP_ALIGN.CENTER, color=WHITE)
    add_textbox(slide, title, x+0.1, 1.9, bw3-0.2, 0.65,
                font_size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_textbox(slide, desc, x+0.1, 2.65, bw3-0.2, 0.95,
                font_size=11, color=MID, align=PP_ALIGN.CENTER)
    if i < len(steps_improve) - 1:
        add_textbox(slide, "→", x+bw3, 2.4, 0.2, 0.4,
                    font_size=14, bold=True, color=I_COL, align=PP_ALIGN.CENTER)

# Two sub-cards
card(slide, 0.5, 4.05, 5.9, 3.15, "Technologie-Stack", [
    "Automation:     n8n (self-hosted / cloud)",
    "AI-Modell:      Claude (Anthropic)",
    "E-Mail:         Gmail (G-Suite)",
    "Dokumente:      Google Docs",
    "Kanäle:         Twilio, WhatsApp, Telegram",
    "Speicher:       n8n Store / Google Sheets",
], phase="I", font_size=12)

card(slide, 6.7, 4.05, 5.9, 3.15, "Schlüsselentscheide", [
    "JSON-Schema erzwingt strukturierten Output",
    "Flat Fields (produkt_1...5) für Kanal-Kompatibilität",
    "Determin. Kalkulation: kein AI-Spielraum",
    "Google Docs statt Plain-Text für Formatierung",
    "Wait-Node statt Zapier HitL (mehr Kontrolle)",
    "Auftragsnr. DR-YYYY-MMDDHHNN als Loop-Anker",
], phase="I", font_size=12)

# ══════════════════════════════════════════════════════════════════
# SLIDE 7 – IMPROVE (Multi-Channel Detail)
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "I", "IMPROVE – Multi-Channel & KI-Analyse", "Alle Kanäle – eine normalisierte Datenbasis")

channels_detail = [
    ("📧 Gmail",        "E-Mail-Thread\nabrufen & zusammenfassen",   "Thread-ID\nals Loop-Anker",   "✅ Aktiv"),
    ("🌐 Webformular",  "Strukturierter\nWebhook-Input",             "Einmalig,\nkein Loop",         "✅ Aktiv"),
    ("📞 Telefon",      "Twilio → Deepgram\nTranskription",          "Aufnahme-URL\n→ Text",         "⏳ Phase 2"),
    ("💬 WhatsApp",     "Business API\n+ Kontext-Speicher",          "Chat-ID\n→ Auftragsnr.",      "⏳ Credentials"),
    ("📋 Formular",     "n8n-Native Form\nstrukturiert",             "Einmalig,\nkein Loop",         "✅ Aktiv"),
    ("✈️ Telegram",     "Bot-Trigger\n+ Kontext-Speicher",           "Chat-ID\n→ Auftragsnr.",      "⏳ Credentials"),
]

for i, (name, desc, note, status) in enumerate(channels_detail):
    row = i // 3
    col_i = i % 3
    x = 0.5 + col_i * 4.25
    y = 1.35 + row * 2.4
    status_col = I_COL if "Aktiv" in status else RGBColor(0xCC, 0x88, 0x00)
    add_rect(slide, x, y, 4.0, 2.1, PHASE_BG["I"], line_color=I_COL)
    add_rect(slide, x, y, 4.0, 0.4, I_COL)
    add_textbox(slide, name, x+0.12, y+0.05, 2.8, 0.32,
                font_size=12, bold=True, color=WHITE)
    add_textbox(slide, status, x+2.95, y+0.05, 1.0, 0.32,
                font_size=10, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    add_textbox(slide, desc, x+0.12, y+0.5, 1.95, 1.35,
                font_size=11, color=DARK)
    add_textbox(slide, note, x+2.1, y+0.5, 1.8, 1.35,
                font_size=11, color=MID, italic=True)

# Bottom merge note
add_rect(slide, 0.5, 6.35, 12.3, 0.85, PHASE_BG["I"], line_color=I_COL)
add_textbox(slide, "→  Alle Kanäle → 'Quellen zusammenführen' → Normalisierter JSON → Claude Vollständigkeitsprüfung",
            0.7, 6.5, 11.8, 0.55, font_size=13, bold=True, color=I_COL)

# ══════════════════════════════════════════════════════════════════
# SLIDE 8 – IMPROVE (Kalkulation & Offerte)
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "I", "IMPROVE – Kalkulation & Offerte", "Deterministischer Preis · Rechtssicheres Dokument · Google Docs")

# Kalkulation
add_rect(slide, 0.5, 1.35, 5.9, 5.85, PHASE_BG["I"], line_color=I_COL)
add_rect(slide, 0.5, 1.35, 5.9, 0.42, I_COL)
add_textbox(slide, "Preiskalkulation (Claude AI – Fixregeln)", 0.65, 1.4,
            5.6, 0.32, font_size=12, bold=True, color=WHITE)
add_multiline(slide, [
    "1.  Gesamtbogen = CEIL((Auflage × 1.08) / 50) × 50",
    "2.  Papierkosten aus Fixpreistabelle (CHF/1000 Bogen)",
    "    90g Offset: 18.– | 170g glänzend: 41.–",
    "    300g Chromo: 68.– | Selbstklebefolie: 95.–",
    "3.  Druckkosten: 55.– Rüst + Bogenrate × Seitenfaktor",
    "4.  Veredelung: Cello 0.065 | Softtouch 0.095 /Bogen",
    "5.  Layout: 95.– / 185.– / 290.– (nur ohne Druckdaten)",
    "6.  Expresszuschlag: knapp +18% | nicht real. +28%",
    "7.  VP netto = Selbstkosten × 1.7241",
    "    → aufrunden auf nächste CHF 5.– | min. CHF 95.–",
    "8.  MwSt. 8.1% | VP brutto = netto + MwSt.",
    "",
    "WICHTIG: Preis = Gesamtpreis ALLER Stück – nicht Stückpreis",
], 0.65, 1.88, 5.6, 5.0, font_size=11, color=DARK)

# Offerte
add_rect(slide, 6.7, 1.35, 5.9, 5.85, PHASE_BG["I"], line_color=I_COL)
add_rect(slide, 6.7, 1.35, 5.9, 0.42, I_COL)
add_textbox(slide, "Angebotsstruktur (Google Docs)", 6.85, 1.4,
            5.6, 0.32, font_size=12, bold=True, color=WHITE)
add_multiline(slide, [
    "1.  Briefkopf (Schellenberg Druck AG, linksbündig)",
    "2.  Empfänger (rechtsbündig)",
    "3.  Angebotsnummer + Datum + Gültig bis (30 Tage)",
    "4.  Betreff (fett)",
    "5.  Leistungsbeschreibung pro Position:",
    "    → Produkt, Format, Farbigkeit, Material",
    "    → Auflage, Veredelung, Druckdaten",
    "    → Kosten: CHF X.- (fett, gerundet)",
    "6.  Kostenaufstellung Pos. 1–X",
    "    Nettobetrag / MwSt. 8.1% / Gesamtbetrag (fett)",
    "7.  Lieferbedingungen",
    "8.  Zahlungsbedingungen (30 Tage / 50% Anzahlung)",
    "9.  Druckdaten-Status",
    "10. Rechtliche Hinweise (Fliesstext, CH-Gerichtsstand)",
    "11. Unterschrift Edgar Glinz",
], 6.85, 1.88, 5.6, 5.0, font_size=11, color=DARK)

# ══════════════════════════════════════════════════════════════════
# SLIDE 9 – CONTROL
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
phase_header(slide, "C", "CONTROL – Kontrolle & Robustheit", "Wie halten wir die Verbesserung stabil?")

control_cards = [
    ("Human in the Loop", [
        "Edgar prüft jedes Angebot vor Versand",
        "Wait-Node pausiert Workflow bis Klick",
        "Approve-Link: ?approved=true",
        "Reject-Link: ?approved=false",
        "Timeout: 72h → Reminder → 48h → Archiv",
        "Execution-ID = Sicherheitstoken",
    ]),
    ("3-Ebenen Fehlerbehandlung", [
        "Ebene 1: Continue on Fail pro Node",
        "         JSON-Validierung nach AI-Steps",
        "         Retry bis 3× bei API-Fehlern",
        "Ebene 2: Wait-Node Timeout-Logik",
        "Ebene 3: Globaler Error-Workflow",
        "         Mail an Edgar bei Absturz",
    ]),
    ("Audit-Trail & Archivierung", [
        "Google Docs: jede Offerte archiviert",
        "Auftrags-Nr. DR-YYYY-MMDDHHNN",
        "Google Sheets: Auftragsdatenbank (geplant)",
        "Schweizer Recht: 10 Jahre Aufbewahrung",
        "Interne Benachrichtigung bei Abschluss",
        "Execution-Log in n8n",
    ]),
    ("Qualitätssicherung", [
        "AI gibt nur valides JSON aus (Schema)",
        "Pflichtfelder erzwungen (niemals leer)",
        "Terminprüfung: realistisch/knapp/unmögl.",
        "Mehrprodukt-Aufträge: Pos. 1–5 geprüft",
        "Kein Ermessensspielraum in Kalkulation",
        "Regelmässige Prompt-Reviews geplant",
    ]),
]

for i, (title, items) in enumerate(control_cards):
    row = i // 2
    col_i = i % 2
    x = 0.5 + col_i * 6.3
    y = 1.35 + row * 2.95
    card(slide, x, y, 6.1, 2.75, title, items, phase="C", font_size=12)

# ══════════════════════════════════════════════════════════════════
# SLIDE 10 – HERAUSFORDERUNGEN (Zapier vs. n8n)
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_rect(slide, 0, 0, 13.33, 0.08, ORANGE)
add_textbox(slide, "Herausforderungen & Lösungen", 0.6, 0.2, 12, 0.65,
            font_size=28, bold=True, color=DARK)
add_textbox(slide, "Lessons Learned aus Phase 1 (Zapier) und Phase 2 (n8n)",
            0.6, 0.88, 12, 0.38, font_size=14, color=MID)

challenges = [
    ("Stückpreis statt Gesamtpreis",
     "AI dividierte VP durch Auflage → Stückpreis",
     "WICHTIG-Hinweis im Prompt: 'Gesamtpreis ALLER Stück'"),
    ("Thread blind",
     "AI las nur neueste E-Mail, ignorierte Kontext",
     "Explizites Thread-Reading + Kurzantwort-Logik"),
    ("Zapier Chip-Fehler",
     "{{...}} als Plaintext getippt → kein Wert",
     "Chip-Picker verwenden, nie manuell tippen"),
    ("Flaches JSON (Zapier)",
     "JSON-Arrays nicht mappbar in Zapier",
     "n8n: Arrays nativ – kein produkt_1...5 Hack nötig"),
    ("HitL immer 'abgelehnt'",
     "Falsche Condition: Status statt Decision-Feld",
     "Field: 'Decision' / Values: approved / rejected"),
    ("Formatierung fehlt",
     "Plain-Text: kein Fett, keine Ausrichtung",
     "n8n + Google Docs: echte Formatierung möglich"),
]

for i, (title, prob, sol) in enumerate(challenges):
    row = i // 2
    col_i = i % 2
    x = 0.5 + col_i * 6.3
    y = 1.45 + row * 1.95
    add_rect(slide, x, y, 6.1, 1.75, LIGHT_BOX, line_color=RGBColor(0xCB, 0xD5, 0xE0))
    add_textbox(slide, title, x+0.15, y+0.08, 5.8, 0.38,
                font_size=13, bold=True, color=DARK)
    add_textbox(slide, f"✗  {prob}", x+0.15, y+0.52, 5.8, 0.4,
                font_size=11, color=RGBColor(0xC0, 0x30, 0x30))
    add_textbox(slide, f"✓  {sol}", x+0.15, y+0.98, 5.8, 0.55,
                font_size=11, color=RGBColor(0x28, 0x72, 0x48))

# ══════════════════════════════════════════════════════════════════
# SLIDE 11 – ERGEBNISSE & ROADMAP
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_rect(slide, 0, 0, 13.33, 0.08, ORANGE)
add_textbox(slide, "Ergebnisse & Roadmap", 0.6, 0.2, 12, 0.65,
            font_size=28, bold=True, color=DARK)
add_textbox(slide, "Was erreicht wurde – was als nächstes kommt",
            0.6, 0.88, 12, 0.38, font_size=14, color=MID)

# Results
add_rect(slide, 0.5, 1.38, 5.9, 5.85, PHASE_BG["I"], line_color=I_COL)
add_rect(slide, 0.5, 1.38, 5.9, 0.42, I_COL)
add_textbox(slide, "Erreicht (Phase 1 & 2)", 0.65, 1.43, 5.6, 0.32,
            font_size=13, bold=True, color=WHITE)
achieved = [
    ("✅", "6 Eingangskanäle normalisiert"),
    ("✅", "Vollständigkeitsprüfung mit Schema"),
    ("✅", "Deterministisches Pricing (0% Spielraum)"),
    ("✅", "Offerte via Google Docs"),
    ("✅", "HitL mit Wait-Node + Webhook-Links"),
    ("✅", "Auftragskontext-Speicher (Loop)"),
    ("✅", "Abschlussmail & Interne Benachrichtigung"),
    ("✅", "DMAIC-Struktur dokumentiert"),
]
for i, (icon, text) in enumerate(achieved):
    add_textbox(slide, f"{icon}  {text}", 0.65, 1.95 + i*0.6,
                5.6, 0.5, font_size=12, color=DARK)

# Roadmap
add_rect(slide, 6.7, 1.38, 5.9, 5.85, PHASE_BG["C"], line_color=C_COL)
add_rect(slide, 6.7, 1.38, 5.9, 0.42, C_COL)
add_textbox(slide, "Roadmap", 6.85, 1.43, 5.6, 0.32,
            font_size=13, bold=True, color=WHITE)

roadmap_items = [
    (ORANGE,          "⚡ Kurzfristig"),
    (MID,   "  →  Fehlerbehandlung (Error-Workflow)"),
    (MID,   "  →  Druckdaten-Anhänge (Google Drive)"),
    (MID,   "  →  End-to-End-Test (Mail + Webhook)"),
    (RGBColor(0xCC,0x88,0x00), "📅 Mittelfristig"),
    (MID,   "  →  Rechnung / CH-Quittung"),
    (MID,   "  →  Auftragsdatenbank (Sheets/Supabase)"),
    (MID,   "  →  Twilio / WhatsApp / Telegram live"),
    (I_COL, "🚀 Langfristig / Produktiv"),
    (MID,   "  →  Telefontranskription (Deepgram)"),
    (MID,   "  →  Kundenstatus-Portal"),
    (MID,   "  →  Produktivbetrieb Schellenberg AG"),
]
y_r = 1.98
for col_r, text in roadmap_items:
    size = 13 if "→" not in text else 11
    bold = "→" not in text
    add_textbox(slide, text, 6.85, y_r, 5.6, 0.42,
                font_size=size, bold=bold, color=col_r)
    y_r += 0.42

# Footer
add_rect(slide, 0, 7.25, 13.33, 0.25, LIGHT_BOX)
add_textbox(slide, "Schellenberg Druck AG · Schützenhausstrasse 5 · 8330 Pfäffikon · schellenbergdruck.ch  |  Juni 2026",
            0.5, 7.28, 12.5, 0.2, font_size=10, color=MID, align=PP_ALIGN.CENTER)

# ── Save ───────────────────────────────────────────────────────
out = "/home/user/zapier-json-schema/KI_Auftragsannahme_DMAIC.pptx"
prs.save(out)
print(f"Saved: {out}")
