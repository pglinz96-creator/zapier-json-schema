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

# ── Save ───────────────────────────────────────────────────────
out = "/home/user/zapier-json-schema/KI_Auftragsannahme_Schellenberg.pptx"
prs.save(out)
print(f"Saved: {out}")
