from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── Brand colours ──────────────────────────────────────────────
DARK_BLUE   = RGBColor(0x1A, 0x2E, 0x4A)   # background / titles
ACCENT      = RGBColor(0xE8, 0x5D, 0x04)   # orange accent
LIGHT       = RGBColor(0xF5, 0xF7, 0xFA)   # light bg boxes
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT   = RGBColor(0x55, 0x65, 0x7A)
GREEN       = RGBColor(0x2D, 0x6A, 0x4F)
YELLOW      = RGBColor(0xF4, 0xA2, 0x61)

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, text, l, t, w, h,
                font_size=18, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
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

def add_rect(slide, l, t, w, h, fill_color, transparency=0):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape

def add_multiline(slide, lines, l, t, w, h, font_size=16, color=WHITE, line_spacing=1.2):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = color

# ══════════════════════════════════════════════════════════════════
# SLIDE 1 – TITLE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)

# Orange accent bar left
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

# Company tag top right
add_textbox(slide, "Schellenberg Druck AG", 1, 0.3, 11.5, 0.5,
            font_size=13, color=YELLOW, align=PP_ALIGN.RIGHT)

# Main title
add_textbox(slide, "KI-gestützte\nAuftragsannahme", 1, 1.8, 10, 2.2,
            font_size=52, bold=True, color=WHITE)

# Subtitle
add_textbox(slide, "Von der Kundenanfrage zum fertigen Angebot – vollautomatisch",
            1, 4.2, 10, 0.8, font_size=22, color=YELLOW)

# Description
add_textbox(slide, "Zapier · Claude AI · Gmail · Human in the Loop",
            1, 5.1, 10, 0.6, font_size=16, color=RGBColor(0xB0, 0xC4, 0xDE))

# Date
add_textbox(slide, "Juni 2026", 1, 6.6, 10, 0.5,
            font_size=13, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════════
# SLIDE 2 – AUSGANGSLAGE & VISION
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Ausgangslage & Vision", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

# Problem box
add_rect(slide, 1, 1.3, 5.3, 4.8, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Vorher", 1.25, 1.4, 4.8, 0.5,
            font_size=18, bold=True, color=ACCENT)
add_multiline(slide, [
    "✗  Jede Anfrage manuell lesen & bewerten",
    "✗  Preise von Hand kalkulieren",
    "✗  Angebote manuell verfassen",
    "✗  Hohes Fehlerrisiko",
    "✗  Zeitaufwand: 30–60 min pro Auftrag",
    "✗  Keine Skalierbarkeit",
], 1.25, 1.95, 4.8, 3.8, font_size=15, color=RGBColor(0xCC, 0xDD, 0xEE))

# Solution box
add_rect(slide, 6.85, 1.3, 5.3, 4.8, RGBColor(0x0D, 0x3B, 0x2A))
add_textbox(slide, "Nachher", 7.1, 1.4, 4.8, 0.5,
            font_size=18, bold=True, color=RGBColor(0x52, 0xB7, 0x88))
add_multiline(slide, [
    "✓  AI liest & analysiert automatisch",
    "✓  Rückfragen automatisch versenden",
    "✓  Preise deterministisch kalkuliert",
    "✓  Angebot KI-generiert",
    "✓  Mensch prüft nur noch & gibt frei",
    "✓  Skalierbar, konsistent, schnell",
], 7.1, 1.95, 4.8, 3.8, font_size=15, color=RGBColor(0xA8, 0xD8, 0xBE))

# Arrow
add_textbox(slide, "→", 6.2, 3.3, 0.7, 0.8, font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 3 – WORKFLOW ÜBERSICHT
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Workflow-Übersicht", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

steps = [
    ("1", "Gmail\nTrigger",         "Neue E-Mail\neingang"),
    ("2", "AI\nAnalyse",            "Daten extrahieren\nThread lesen"),
    ("3", "Path\nEntscheid",        "Vollständig?\nJa / Nein"),
    ("6", "Kalkulation\nAI",        "Fixpreise\nberechnen"),
    ("7", "Offerte\nAI",            "Angebot\ngenerieren"),
    ("9", "Human in\nthe Loop",     "Edgar prüft\n& gibt frei"),
]

colors_step = [
    RGBColor(0x2B, 0x4C, 0x7E),
    RGBColor(0x2B, 0x4C, 0x7E),
    RGBColor(0x6B, 0x35, 0x10),
    RGBColor(0x1A, 0x53, 0x3A),
    RGBColor(0x1A, 0x53, 0x3A),
    RGBColor(0x5C, 0x20, 0x6A),
]

box_w = 1.7
gap   = 0.25
start = 0.9
for i, (num, title, desc) in enumerate(steps):
    x = start + i * (box_w + gap)
    add_rect(slide, x, 1.5, box_w, 3.8, colors_step[i])
    add_textbox(slide, f"Step {num}", x+0.1, 1.6, box_w-0.2, 0.4,
                font_size=11, color=YELLOW)
    add_textbox(slide, title, x+0.1, 2.0, box_w-0.2, 0.9,
                font_size=15, bold=True, color=WHITE)
    add_textbox(slide, desc, x+0.1, 3.0, box_w-0.2, 1.5,
                font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE))
    if i < len(steps) - 1:
        add_textbox(slide, "→", x + box_w, 2.7, gap + 0.1, 0.5,
                    font_size=20, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# Fork note
add_rect(slide, 0.9, 5.6, 5, 1.1, RGBColor(0x6B, 0x35, 0x10))
add_textbox(slide, "Unvollständig → Automatische Rückfrage an Kunden (Step 16)",
            1.05, 5.7, 4.7, 0.8, font_size=13, color=YELLOW)

add_rect(slide, 6.2, 5.6, 5.9, 1.1, RGBColor(0x1A, 0x53, 0x3A))
add_textbox(slide, "Vollständig → Kalkulation → Offerte → Freigabe → Versand",
            6.35, 5.7, 5.6, 0.8, font_size=13, color=RGBColor(0x90, 0xE0, 0xB8))

# ══════════════════════════════════════════════════════════════════
# SLIDE 4 – AI ANALYSE PROMPT
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Step 2 – AI Auftragsanalyse", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Claude analysiert den gesamten E-Mail-Thread und extrahiert alle Auftragsdaten",
            1, 1.1, 11, 0.5, font_size=15, color=YELLOW)

# Left col
add_rect(slide, 1, 1.8, 5.3, 4.9, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Pflichtfelder", 1.2, 1.9, 4.9, 0.45,
            font_size=16, bold=True, color=ACCENT)
add_multiline(slide, [
    "Kundendaten: Name, Firma, Adresse, Telefon, E-Mail",
    "",
    "Auftragsdaten:",
    "· Produkt & Format",
    "· Auflage (Stückzahl)",
    "· Farbigkeit (z.B. 4/0 = einseitig)",
    "· Material & Grammatur",
    "· Veredelung (oder keine)",
    "· Liefertermin",
    "· Druckdaten vorhanden?",
], 1.2, 2.45, 4.9, 3.9, font_size=13, color=RGBColor(0xB0, 0xC4, 0xDE))

# Right col
add_rect(slide, 6.85, 1.8, 5.3, 4.9, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Intelligente Regeln", 7.05, 1.9, 4.9, 0.45,
            font_size=16, bold=True, color=RGBColor(0x52, 0xB7, 0x88))
add_multiline(slide, [
    "Thread-Auswertung: alle Mails im Verlauf",
    "→ bei Widerspruch gilt neuere Angabe",
    "",
    "Keine Doppelfragen:",
    "· 4/0 = einseitig (nicht nochmal fragen)",
    "· 4/4 = beidseitig (klar)",
    "· 'keine Veredelung' = direkt eintragen",
    "",
    "E-Mail autom. aus Absender-Feld",
    "",
    "Ton: 1. Mail = freundlich | Folge = kurz",
], 7.05, 2.45, 4.9, 3.9, font_size=13, color=RGBColor(0xA8, 0xD8, 0xBE))

# ══════════════════════════════════════════════════════════════════
# SLIDE 5 – KALKULATION
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Step 6 – Preiskalkulation", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Deterministisch · Fixwerte · Keine Ermessensspielräume",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Formula box
add_rect(slide, 1, 1.65, 11.1, 1.55, RGBColor(0x0D, 0x1F, 0x33))
add_textbox(slide, "Gesamtbogen = AUFRUNDEN((Auflage × 1.08) / 50) × 50   |   VP netto = Selbstkosten × 1.7241   →   aufrunden auf CHF 5.–",
            1.2, 1.75, 10.7, 0.55, font_size=13, bold=True, color=ACCENT)
add_textbox(slide, "Selbstkosten = Papier + Druck + Veredelung + Layout + Expresszuschlag   |   MwSt. 8.1%   |   Min. CHF 95.–",
            1.2, 2.3, 10.7, 0.55, font_size=13, color=RGBColor(0xB0, 0xC4, 0xDE))

# Cost blocks
blocks = [
    ("Papierkosten", "CHF/1000 Bogen\n90g Offset: 18.–\n135g matt: 32.–\n170g glänzend: 41.–\n300g Chromo: 68.–", RGBColor(0x1A, 0x3A, 0x5C)),
    ("Druckkosten", "Rüst: CHF 55.–\n≤1000: 0.048/Bogen\n1001–5000: 0.034\n>5000: 0.024\nBeidseitig: ×1.75", RGBColor(0x1A, 0x3A, 0x5C)),
    ("Veredelung", "Cello eins.: 0.065\nSofttouch: 0.095\nUV vollfl.: 0.075\nHeissfolie: 0.185\nStanzung: 0.085+190", RGBColor(0x1A, 0x3A, 0x5C)),
    ("Marge", "Handelsmarge:\n42% auf VP netto\n= Aufschlag\n72.41% auf SK\nMin. CHF 95.–", RGBColor(0x2D, 0x5A, 0x27)),
]

bw = 2.55
for i, (title, content, col) in enumerate(blocks):
    x = 1 + i * (bw + 0.13)
    add_rect(slide, x, 3.4, bw, 3.7, col)
    add_textbox(slide, title, x+0.15, 3.5, bw-0.3, 0.5,
                font_size=14, bold=True, color=ACCENT)
    add_textbox(slide, content, x+0.15, 4.05, bw-0.3, 2.9,
                font_size=12, color=WHITE)

# ══════════════════════════════════════════════════════════════════
# SLIDE 6 – OFFERTE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Step 7 – Angebotserstellung", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Claude erstellt ein rechtssicheres Schweizer Geschäftsdokument – vollautomatisch",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Document preview box
add_rect(slide, 1, 1.65, 6.2, 5.5, RGBColor(0xF5, 0xF5, 0xF0))
add_textbox(slide, "Schellenberg Druck AG\nSchützenhausstrasse 5 · 8330 Pfäffikon\nTelefon: 044 953 11 11",
            1.2, 1.8, 5.8, 0.8, font_size=11, color=DARK_BLUE)
add_textbox(slide, "                                    Max Mustermann\n                                    Musterstrasse 1\n                                    8000 Zürich",
            1.2, 2.65, 5.8, 0.7, font_size=11, color=DARK_BLUE)
add_textbox(slide, "Angebot Nr. DR-2026-47382 – Flyer A5",
            1.2, 3.5, 5.8, 0.4, font_size=12, bold=True, color=DARK_BLUE)
add_textbox(slide, "1. Leistungsbeschreibung\nPosition 1: Flyer A5, 4/0, 135g matt, 1000 Stk\nNettobetrag exkl. MwSt.:  CHF 285.00\nMwSt. 8.1%:                  CHF 23.09\nGesamtbetrag inkl. MwSt.: CHF 308.09\n\n2. Zahlungsbedingungen\n30 Tage netto\n\n3. Rechtliche Hinweise\nAngebot gültig bis ...",
            1.2, 4.0, 5.8, 2.9, font_size=10, color=RGBColor(0x33, 0x33, 0x33))

# Rules box
add_rect(slide, 7.5, 1.65, 4.65, 5.5, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Layout-Regeln", 7.7, 1.75, 4.2, 0.45,
            font_size=15, bold=True, color=ACCENT)
add_multiline(slide, [
    "✓ Kein Markdown (**,---,###)",
    "✓ Kein Tabellen oder Trennlinien",
    "✓ Empfänger rechtsbündig",
    "✓ Titel fett, normale Gross-/Klein-",
    "   schreibung",
    "✓ Max. 1 Leerzeile zwischen Abschnitten",
    "✓ Max. 2 Seiten",
    "✓ MwSt. 8.1% immer separat",
    "✓ Preise aus Step 6 direkt eingesetzt",
    "✓ Unterschrift: Edgar Glinz",
    "",
    "Input aus Step 6:",
    "· verkaufspreis_netto",
    "· mwst_betrag",
    "· verkaufspreis_brutto",
    "· preis_pro_stueck_netto",
], 7.7, 2.3, 4.2, 4.6, font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE))

# ══════════════════════════════════════════════════════════════════
# SLIDE 7 – HUMAN IN THE LOOP
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Step 9 – Human in the Loop", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Der Mensch bleibt Entscheider – die Maschine bereitet vor",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Flow
stages = [
    ("📧", "Angebot\nbereit", "KI hat Offerte\ngeneriert", DARK_BLUE),
    ("🔔", "Edgar\nerhält Mail", "Content to review:\nformatiertes Angebot", RGBColor(0x2B, 0x4C, 0x7E)),
    ("✅", "Freigabe\nerteilt", "→ Angebot an\n   Kunden senden", RGBColor(0x1A, 0x53, 0x3A)),
    ("❌", "Abgelehnt", "→ Ablehnungs-Mail\n   mit Begründung", RGBColor(0x7A, 0x1A, 0x1A)),
]

for i, (icon, title, desc, col) in enumerate(stages):
    x = 1 + i * 2.9
    add_rect(slide, x, 1.8, 2.6, 3.5, col)
    add_textbox(slide, icon, x+0.1, 1.9, 2.4, 0.7, font_size=28, align=PP_ALIGN.CENTER)
    add_textbox(slide, title, x+0.1, 2.65, 2.4, 0.6, font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, desc, x+0.1, 3.35, 2.4, 1.7, font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE), align=PP_ALIGN.CENTER)
    if i < 2:
        add_textbox(slide, "→", x + 2.6, 3.0, 0.3, 0.5, font_size=22, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# Config note
add_rect(slide, 1, 5.5, 11.1, 1.65, RGBColor(0x0D, 0x1F, 0x33))
add_textbox(slide, "Konfiguration:", 1.2, 5.6, 3, 0.4, font_size=13, bold=True, color=ACCENT)
add_textbox(slide, "Approver: p.glinz96@gmail.com   ·   Content to review: Step 7 Output (Offerte-Text)   ·   Path-Bedingung: Status = Approved / Rejected (exakt)",
            1.2, 6.05, 10.7, 0.8, font_size=12, color=WHITE)

# ══════════════════════════════════════════════════════════════════
# SLIDE 8 – HERAUSFORDERUNGEN & LÖSUNGEN
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Herausforderungen & Lösungen", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

challenges = [
    ("Zufällige Preise",      "AI kalkulierte jeden Run anders",           "Fixe Preistabellen, kein Ermessensspielraum"),
    ("Doppelte Rückfragen",   "'4/0 = einseitig' wurde nochmals gefragt",  "Interpretationsregeln im Prompt"),
    ("Thread blind",          "AI las nur neueste Mail",                    "Explizites Thread-Reading, Verlauf auswerten"),
    ("Zapier Chip-Fehler",    "{{...}} als Text → kein Wert erkannt",       "Chip-Picker statt manueller Eingabe"),
    ("Preise fehlen Offerte", "Step 7 kannte Step 6 Werte nicht",           "Input Fields mit 4 Preis-Chips aus Step 6"),
    ("HitL zeigt falsch",     "Kunden-Mail statt Angebot zur Freigabe",    "Content to review → Step 7 Output"),
]

for i, (prob_title, prob, sol) in enumerate(challenges):
    row = i // 2
    col = i % 2
    x = 1 + col * 6.2
    y = 1.3 + row * 1.95
    add_rect(slide, x, y, 5.9, 1.75, RGBColor(0x12, 0x22, 0x38))
    add_textbox(slide, prob_title, x+0.15, y+0.08, 5.6, 0.4,
                font_size=13, bold=True, color=ACCENT)
    add_textbox(slide, f"✗ {prob}", x+0.15, y+0.5, 5.6, 0.4,
                font_size=11, color=RGBColor(0xFF, 0x9A, 0x9A))
    add_textbox(slide, f"✓ {sol}", x+0.15, y+0.95, 5.6, 0.55,
                font_size=11, color=RGBColor(0x90, 0xE0, 0xB8))

# ══════════════════════════════════════════════════════════════════
# SLIDE 9 – LEKTIONEN & ERKENNTNISSE
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Key Learnings", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

learnings = [
    ("01", "Fixwerte statt Ranges",       "AI-Kalkulationen brauchen exakte Preistabellen – keine Spielräume."),
    ("02", "Thread-Kontext ist alles",    "Der gesamte E-Mail-Verlauf muss ausgewertet werden, nicht nur die letzte Mail."),
    ("03", "Flache JSON-Struktur",        "Zapier braucht Top-Level-Felder – verschachtelte Objekte können nicht direkt gemappt werden."),
    ("04", "Chips ≠ Plaintext",           "In Zapier niemals {{...}} manuell tippen – immer den Chip-Picker verwenden."),
    ("05", "Input Fields explizit",       "Jeder AI-Step braucht jeden einzelnen Input-Wert als gemappten Chip."),
    ("06", "HitL nur live testbar",       "Approve/Reject funktioniert nur bei veröffentlichtem Zap mit echtem Trigger."),
]

for i, (num, title, desc) in enumerate(learnings):
    row = i // 2
    col_idx = i % 2
    x = 1 + col_idx * 6.1
    y = 1.3 + row * 1.95
    add_rect(slide, x, y, 5.8, 1.75, RGBColor(0x12, 0x22, 0x38))
    add_textbox(slide, num, x+0.15, y+0.1, 0.7, 0.5, font_size=22, bold=True, color=ACCENT)
    add_textbox(slide, title, x+0.9, y+0.1, 4.7, 0.5, font_size=14, bold=True, color=WHITE)
    add_textbox(slide, desc, x+0.9, y+0.65, 4.7, 0.95, font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE))

# ══════════════════════════════════════════════════════════════════
# SLIDE 10 – NÄCHSTE SCHRITTE & AUSBLICK
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Status & Nächste Schritte", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

# Status column
add_rect(slide, 1, 1.3, 5.3, 5.8, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Aktueller Status", 1.2, 1.4, 4.9, 0.5,
            font_size=16, bold=True, color=ACCENT)
status_items = [
    ("✅", "Gmail-Trigger", "Live"),
    ("✅", "AI Analyse (Step 2)", "Thread-aware"),
    ("✅", "Paths Vollst./Unvollst.", "Konfiguriert"),
    ("✅", "Kalkulation (Step 6)", "Fixpreise, deterministisch"),
    ("✅", "Offerte (Step 7)", "Mit Kalkulations-Preisen"),
    ("✅", "Human in the Loop", "Approval-Mail aktiv"),
    ("⚠️", "Path nach HitL", "Condition prüfen"),
    ("⚠️", "Rückfrage-Mail To-Feld", "Chip umstellen"),
]
for i, (icon, item, note) in enumerate(status_items):
    y = 2.0 + i * 0.6
    add_textbox(slide, f"{icon}  {item}", 1.2, y, 3.2, 0.5, font_size=12, color=WHITE)
    add_textbox(slide, note, 4.4, y, 1.8, 0.5, font_size=11, color=GREY_TEXT)

# Next steps column
add_rect(slide, 6.85, 1.3, 5.3, 5.8, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Nächste Schritte", 7.05, 1.4, 4.9, 0.5,
            font_size=16, bold=True, color=RGBColor(0x52, 0xB7, 0x88))
next_steps = [
    "→ To-Feld Step 16: Chip auf kunde_email",
    "→ Path Step 11/13: Approved / Rejected prüfen",
    "→ Live-Test: vollständiger Auftrag",
    "→ HitL Approve → Angebot-Versand testen",
    "→ Google Docs: PDF-Anhang für Kunden",
    "→ Multi-Produkt-Aufträge validieren",
    "→ Sonderprodukte ergänzen (Pappbecher etc.)",
]
for i, step in enumerate(next_steps):
    y = 2.0 + i * 0.7
    add_textbox(slide, step, 7.05, y, 4.9, 0.6, font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE))

# Vision footer
add_rect(slide, 1, 7.1, 11.1, 0.2, ACCENT)
add_textbox(slide, "Schellenberg Druck AG · Schützenhausstrasse 5 · 8330 Pfäffikon · schellenbergdruck.ch",
            1, 7.1, 11.1, 0.35, font_size=10, color=GREY_TEXT, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 11 – VON ZAPIER ZU N8N
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Von Zapier zu n8n", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Phase 2 – Plattformwechsel & Architekturentscheid",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Zapier column
add_rect(slide, 1, 1.65, 5.3, 5.0, RGBColor(0x3A, 0x1A, 0x10))
add_textbox(slide, "Zapier – Limitierungen", 1.2, 1.75, 4.9, 0.45,
            font_size=16, bold=True, color=RGBColor(0xFF, 0x7A, 0x5A))
add_multiline(slide, [
    "✗  JSON-Arrays nicht nativ unterstützt",
    "✗  Flat Fields nötig (produkt_1...5)",
    "✗  Formatter für Textformatierung unzureichend",
    "✗  HitL: Decision-Feld 'approved'/'rejected'",
    "   oft falsch konfiguriert",
    "✗  Hohe Kosten pro Task/Run",
    "✗  Keine Code-Nodes (kein Python/JS)",
    "✗  Begrenzte Debugging-Möglichkeiten",
], 1.2, 2.3, 4.9, 4.0, font_size=13, color=RGBColor(0xFF, 0xCC, 0xBB))

# Arrow
add_textbox(slide, "→", 6.2, 3.7, 0.8, 0.8,
            font_size=36, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# n8n column
add_rect(slide, 7.05, 1.65, 5.3, 5.0, RGBColor(0x0D, 0x3B, 0x2A))
add_textbox(slide, "n8n – Vorteile", 7.25, 1.75, 4.9, 0.45,
            font_size=16, bold=True, color=RGBColor(0x52, 0xB7, 0x88))
add_multiline(slide, [
    "✓  JSON-Arrays nativ – kein Flat-Field-Hack",
    "✓  Code-Node (JS/Python) für jede Logik",
    "✓  Wait-Node: echter HitL mit Webhook",
    "✓  Google Docs für formatierte Offerten",
    "✓  Self-hosted oder Cloud – günstig",
    "✓  Vollständiges Error-Handling möglich",
    "✓  Sauberes Debugging & Execution-Log",
    "✓  Kein Task-Limit",
], 7.25, 2.3, 4.9, 4.0, font_size=13, color=RGBColor(0xA8, 0xD8, 0xBE))

# ══════════════════════════════════════════════════════════════════
# SLIDE 12 – MULTI-CHANNEL ARCHITEKTUR
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Multi-Channel Eingangsarchitektur", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "6 Kanäle – ein normalisierter Datenstrom – eine KI-Analyse",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

channels = [
    ("📧", "Gmail",          "E-Mail & Thread\nabrufen",           RGBColor(0x1A, 0x3A, 0x6C)),
    ("🌐", "Webformular",    "Online-Bestellung\nWebhook POST",    RGBColor(0x1A, 0x3A, 0x6C)),
    ("📞", "Telefon",        "Twilio → Transkript\n(Phase 2)",     RGBColor(0x4A, 0x2A, 0x0A)),
    ("💬", "WhatsApp",       "Business API\n+ Kontext-Store",      RGBColor(0x0D, 0x3B, 0x1A)),
    ("📋", "Kontaktformular","n8n-Form\nstrukturiert",             RGBColor(0x1A, 0x3A, 0x6C)),
    ("✈️", "Telegram",       "Bot-Trigger\n+ Kontext-Store",       RGBColor(0x0D, 0x2A, 0x4A)),
]

bw, bh = 1.85, 2.5
for i, (icon, name, desc, col) in enumerate(channels):
    x = 0.7 + i * (bw + 0.22)
    add_rect(slide, x, 1.65, bw, bh, col)
    add_textbox(slide, icon, x, 1.75, bw, 0.65, font_size=26, align=PP_ALIGN.CENTER)
    add_textbox(slide, name, x+0.1, 2.45, bw-0.2, 0.45,
                font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, desc, x+0.1, 2.95, bw-0.2, 1.1,
                font_size=11, color=RGBColor(0xB0, 0xC4, 0xDE), align=PP_ALIGN.CENTER)
    # down arrow
    add_textbox(slide, "↓", x + bw/2 - 0.15, 4.25, 0.4, 0.4,
                font_size=18, color=ACCENT, align=PP_ALIGN.CENTER)

# Merge node
add_rect(slide, 3.5, 4.75, 5.8, 0.8, RGBColor(0x2B, 0x4C, 0x7E))
add_textbox(slide, "Quellen zusammenführen  →  Normalisierter JSON-Input",
            3.65, 4.85, 5.5, 0.55, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_textbox(slide, "↓", 6.2, 5.65, 0.5, 0.4, font_size=18, color=ACCENT, align=PP_ALIGN.CENTER)

# Analysis node
add_rect(slide, 3.5, 6.1, 5.8, 0.85, RGBColor(0x1A, 0x53, 0x3A))
add_textbox(slide, "Vollständigkeitsprüfung (Claude AI)  →  Auftrag vollständig?",
            3.65, 6.2, 5.5, 0.55, font_size=14, bold=True, color=RGBColor(0x90, 0xE0, 0xB8), align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 13 – N8N WORKFLOW ARCHITEKTUR
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "n8n Workflow – Vollständige Architektur", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)

# Complete path
add_rect(slide, 1, 1.2, 11.1, 0.4, RGBColor(0x1A, 0x53, 0x3A))
add_textbox(slide, "VOLLSTÄNDIG", 1.15, 1.25, 2, 0.3, font_size=11, bold=True, color=RGBColor(0x90, 0xE0, 0xB8))

complete_steps = [
    "Eingangs-\nbestätigung",
    "Kalkulation\n(Claude AI)",
    "Google Docs\nOfferte",
    "Zur Freigabe\nsenden",
    "Wait-Node\n(HitL)",
    "Entscheidung\nprüfen (IF)",
    "Genehmigt\n→ Versand",
]
cw = 1.4
for i, s in enumerate(complete_steps):
    x = 1 + i * (cw + 0.15)
    add_rect(slide, x, 1.7, cw, 1.5, RGBColor(0x0D, 0x3B, 0x2A))
    add_textbox(slide, s, x+0.05, 1.8, cw-0.1, 1.2, font_size=10, color=WHITE, align=PP_ALIGN.CENTER)
    if i < len(complete_steps) - 1:
        add_textbox(slide, "→", x+cw, 2.25, 0.18, 0.4, font_size=12, color=ACCENT, align=PP_ALIGN.CENTER)

# Incomplete path
add_rect(slide, 1, 3.4, 11.1, 0.4, RGBColor(0x6B, 0x35, 0x10))
add_textbox(slide, "UNVOLLSTÄNDIG", 1.15, 3.45, 2.5, 0.3, font_size=11, bold=True, color=YELLOW)

incomplete_steps = [
    "Nachfrage-\nEmail generieren",
    "Auftragskontext\nspeichern",
    "Nachfrage\nsenden",
    "Kunde antwortet\n(loop back)",
]
for i, s in enumerate(incomplete_steps):
    x = 1 + i * (cw + 0.5)
    add_rect(slide, x, 3.9, cw+0.3, 1.5, RGBColor(0x4A, 0x22, 0x08))
    add_textbox(slide, s, x+0.05, 4.0, cw+0.2, 1.2, font_size=10, color=YELLOW, align=PP_ALIGN.CENTER)
    if i < len(incomplete_steps) - 1:
        add_textbox(slide, "→", x+cw+0.3, 4.45, 0.55, 0.4, font_size=12, color=ACCENT, align=PP_ALIGN.CENTER)

# Error path
add_rect(slide, 1, 5.6, 11.1, 0.4, RGBColor(0x5C, 0x1A, 0x1A))
add_textbox(slide, "FEHLER", 1.15, 5.65, 1.5, 0.3, font_size=11, bold=True, color=RGBColor(0xFF, 0x9A, 0x9A))

error_steps = [
    "Error-Trigger\n(global)",
    "Fehler\nklassifizieren",
    "Benachrichtigung\nan Edgar",
    "Execution-Log\nspeichern",
]
for i, s in enumerate(error_steps):
    x = 1 + i * (cw + 0.5)
    add_rect(slide, x, 6.1, cw+0.3, 1.1, RGBColor(0x40, 0x10, 0x10))
    add_textbox(slide, s, x+0.05, 6.18, cw+0.2, 0.9, font_size=10, color=RGBColor(0xFF, 0xCC, 0xCC), align=PP_ALIGN.CENTER)
    if i < len(error_steps) - 1:
        add_textbox(slide, "→", x+cw+0.3, 6.5, 0.55, 0.4, font_size=12, color=ACCENT, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 14 – N8N HUMAN IN THE LOOP
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Human in the Loop – n8n Wait-Node", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Workflow pausiert – Edgar entscheidet – Webhook setzt fort",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Flow boxes
hitl_steps = [
    ("1", "Offerte\nfertig", "Google Docs\ngeneriert", RGBColor(0x1A, 0x3A, 0x6C)),
    ("2", "E-Mail\nan Edgar", "Offerte-Link +\nApprove/Reject", RGBColor(0x2B, 0x4C, 0x7E)),
    ("3", "Wait-Node\npausiert", "Wartet auf\nWebhook-Klick", RGBColor(0x5C, 0x20, 0x6A)),
    ("4", "Edgar\nklickt Link", "?approved=true\noder false", RGBColor(0x1A, 0x53, 0x3A)),
    ("5", "IF-Node\nentscheidet", "true → Versand\nfalse → Absage", RGBColor(0x1A, 0x53, 0x3A)),
]
bw2 = 2.1
for i, (num, title, desc, col) in enumerate(hitl_steps):
    x = 0.9 + i * (bw2 + 0.22)
    add_rect(slide, x, 1.65, bw2, 3.0, col)
    add_textbox(slide, num, x+0.1, 1.75, 0.5, 0.45, font_size=18, bold=True, color=ACCENT)
    add_textbox(slide, title, x+0.1, 2.25, bw2-0.2, 0.7,
                font_size=14, bold=True, color=WHITE)
    add_textbox(slide, desc, x+0.1, 3.0, bw2-0.2, 1.4,
                font_size=12, color=RGBColor(0xB0, 0xC4, 0xDE))
    if i < len(hitl_steps) - 1:
        add_textbox(slide, "→", x+bw2, 2.85, 0.25, 0.5,
                    font_size=18, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

# URL example
add_rect(slide, 1, 4.85, 11.1, 1.2, RGBColor(0x0D, 0x1F, 0x33))
add_textbox(slide, "Approve-Link:", 1.2, 4.95, 2, 0.35, font_size=12, bold=True, color=ACCENT)
add_textbox(slide, "{{ $execution.resumeUrl }}?approved=true",
            3.3, 4.95, 8.5, 0.35, font_size=12, color=RGBColor(0x90, 0xE0, 0xB8))
add_textbox(slide, "Reject-Link:", 1.2, 5.45, 2, 0.35, font_size=12, bold=True, color=RGBColor(0xFF, 0x7A, 0x5A))
add_textbox(slide, "{{ $execution.resumeUrl }}?approved=false",
            3.3, 5.45, 8.5, 0.35, font_size=12, color=RGBColor(0xFF, 0xBB, 0xAA))

# Timeout note
add_rect(slide, 1, 6.2, 11.1, 1.0, RGBColor(0x1A, 0x2A, 0x45))
add_textbox(slide, "⏰  Timeout-Logik:", 1.2, 6.3, 2.5, 0.35, font_size=12, bold=True, color=YELLOW)
add_textbox(slide, "72h kein Klick → Reminder-Mail an Edgar  →  48h  →  Auftrag automatisch archiviert",
            3.8, 6.3, 8, 0.35, font_size=12, color=RGBColor(0xF4, 0xA2, 0x61))
add_textbox(slide, "Execution-ID in URL = Sicherheitstoken – kein unauthorized Trigger möglich",
            1.2, 6.72, 10.7, 0.35, font_size=11, italic=True, color=GREY_TEXT)

# ══════════════════════════════════════════════════════════════════
# SLIDE 15 – FEHLERBEHANDLUNG
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Fehlerbehandlung & Robustheit", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "3-Ebenen-Konzept für produktionssicheren Betrieb",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

levels = [
    (
        "Ebene 1 – Node-Level",
        RGBColor(0x1A, 0x3A, 0x6C),
        [
            "Continue on Fail = ON bei allen kritischen Nodes",
            "AI-Steps: JSON-Validierung nach jedem Output",
            "Gmail/Google Docs: Retry bis 3× bei Fehler",
            "IF-Node: {{ $json.auftrag_nr }} vorhanden?",
        ]
    ),
    (
        "Ebene 2 – Wait-Node Timeout",
        RGBColor(0x3A, 0x1A, 0x6C),
        [
            "Maximale Wartezeit: 72 Stunden",
            "Nach Ablauf: Reminder-Mail an Edgar",
            "Nach 48h ohne Reaktion: Auto-Archivierung",
            "Kunde erhält Status-Update",
        ]
    ),
    (
        "Ebene 3 – Globaler Error-Workflow",
        RGBColor(0x5C, 0x1A, 0x1A),
        [
            "Separater Error-Workflow (n8n Settings)",
            "Fängt ALLE unbehandelten Fehler aller Workflows",
            "Mail an Edgar: Workflow + Node + Fehlermeldung",
            "Link zur Execution für direktes Debugging",
        ]
    ),
]

for i, (title, col, items) in enumerate(levels):
    x = 1 + i * 3.85
    add_rect(slide, x, 1.65, 3.6, 5.1, col)
    add_textbox(slide, title, x+0.15, 1.75, 3.3, 0.5,
                font_size=13, bold=True, color=WHITE)
    for j, item in enumerate(items):
        add_textbox(slide, f"→  {item}", x+0.15, 2.4 + j*0.85, 3.3, 0.75,
                    font_size=11, color=RGBColor(0xCC, 0xDD, 0xFF))

# Critical failure points
add_rect(slide, 1, 7.0, 11.1, 0.35, RGBColor(0x0D, 0x1F, 0x33))
add_textbox(slide, "Kritische Punkte: Anthropic Timeout · Gmail Auth · Google Docs API-Limit · Ungültiger JSON-Output der AI",
            1.2, 7.05, 10.7, 0.25, font_size=11, color=GREY_TEXT, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
# SLIDE 16 – AKTUELLER STAND & ROADMAP
# ══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DARK_BLUE)
add_rect(slide, 0, 0, 0.45, 7.5, ACCENT)

add_textbox(slide, "Aktueller Stand & Roadmap", 1, 0.35, 11, 0.7,
            font_size=30, bold=True, color=WHITE)
add_textbox(slide, "Juni 2026 – Phase 2 (n8n) in Entwicklung",
            1, 1.05, 11, 0.45, font_size=15, color=YELLOW)

# Done column
add_rect(slide, 1, 1.65, 5.3, 5.5, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Implementiert", 1.2, 1.75, 4.9, 0.45,
            font_size=16, bold=True, color=RGBColor(0x52, 0xB7, 0x88))
done_items = [
    "✅  6 Eingangskanäle (Mail, Web, WA, TG, Tel, Form)",
    "✅  Vollständigkeitsprüfung mit Schema",
    "✅  Preiskalkulation (deterministisch)",
    "✅  Offerte via Google Docs",
    "✅  HitL mit Wait-Node + Webhook-Links",
    "✅  Auftragskontext-Speicher (Loop-Handling)",
    "✅  Abschlussmail & Interne Benachrichtigung",
    "✅  Ablehungs-Pfad konzeptioniert",
]
for i, item in enumerate(done_items):
    add_textbox(slide, item, 1.2, 2.3 + i*0.58, 4.9, 0.5,
                font_size=12, color=RGBColor(0xA8, 0xD8, 0xBE))

# Roadmap column
add_rect(slide, 6.85, 1.65, 5.3, 5.5, RGBColor(0x12, 0x22, 0x38))
add_textbox(slide, "Roadmap", 7.05, 1.75, 4.9, 0.45,
            font_size=16, bold=True, color=ACCENT)
roadmap = [
    ("⚡ Kurzfristig",  RGBColor(0xFF, 0xCC, 0x00), [
        "Fehlerbehandlung (Error-Workflow)",
        "Druckdaten-Anhänge (Drive)",
        "End-to-End-Test Gmail + Webhook",
    ]),
    ("📅 Mittelfristig", RGBColor(0xF4, 0xA2, 0x61), [
        "Rechnung / CH-Quittung generieren",
        "Auftragsdatenbank (Sheets/Supabase)",
        "Twilio/WhatsApp Credentials",
    ]),
    ("🚀 Langfristig",  RGBColor(0x90, 0xE0, 0xB8), [
        "Telefontranskription (Deepgram)",
        "Kundenstatus-Portal",
        "Produktivbetrieb Schellenberg AG",
    ]),
]
y_off = 2.3
for emoji_title, col, items in roadmap:
    add_textbox(slide, emoji_title, 7.05, y_off, 4.9, 0.38,
                font_size=12, bold=True, color=col)
    y_off += 0.42
    for item in items:
        add_textbox(slide, f"  →  {item}", 7.05, y_off, 4.9, 0.38,
                    font_size=11, color=RGBColor(0xB0, 0xC4, 0xDE))
        y_off += 0.38
    y_off += 0.15

# Footer
add_rect(slide, 1, 7.1, 11.1, 0.2, ACCENT)
add_textbox(slide, "Schellenberg Druck AG · Schützenhausstrasse 5 · 8330 Pfäffikon · schellenbergdruck.ch",
            1, 7.1, 11.1, 0.35, font_size=10, color=GREY_TEXT, align=PP_ALIGN.CENTER)

# ── Save ───────────────────────────────────────────────────────
out = "/home/user/zapier-json-schema/KI_Auftragsannahme_Schellenberg.pptx"
prs.save(out)
print(f"Saved: {out}")
