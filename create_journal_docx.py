from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ───────────────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(3.0)
section.right_margin  = Cm(2.5)

# ── Colours ────────────────────────────────────────────────────
DARK_BLUE = RGBColor(0x1A, 0x2E, 0x4A)
ORANGE    = RGBColor(0xE8, 0x5D, 0x04)
GREY      = RGBColor(0x55, 0x65, 0x7A)
GREEN     = RGBColor(0x2D, 0x6A, 0x4F)
RED_SOFT  = RGBColor(0xC0, 0x39, 0x2B)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

def set_run_font(run, size=11, bold=False, italic=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18 if level == 1 else 12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, size=18, bold=True, color=DARK_BLUE)
    elif level == 2:
        set_run_font(run, size=14, bold=True, color=ORANGE)
    elif level == 3:
        set_run_font(run, size=12, bold=True, color=DARK_BLUE)
    return p

def add_body(doc, text, color=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    set_run_font(run, size=11, italic=italic, color=color or RGBColor(0x1A, 0x1A, 0x1A))
    return p

def add_bullet(doc, text, bold_prefix=None, color=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.8)
    if bold_prefix:
        r1 = p.add_run(bold_prefix + " ")
        set_run_font(r1, size=11, bold=True, color=ORANGE)
    r2 = p.add_run(text)
    set_run_font(r2, size=11, color=color or RGBColor(0x1A, 0x1A, 0x1A))
    return p

def add_code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Cm(1.0)
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), 'F0F0F0')
    p._p.get_or_add_pPr().append(shading)
    run = p.add_run(text)
    run.font.name = 'Courier New'
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
    return p

def add_divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'E85D04')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

def add_status_table(doc):
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    for cell, txt in zip(hdr, ['Komponente', 'Status', 'Details']):
        cell.text = txt
        run = cell.paragraphs[0].runs[0]
        set_run_font(run, size=10, bold=True, color=WHITE)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1A2E4A')
        cell._tc.get_or_add_tcPr().append(shd)

    rows = [
        ('Gmail-Trigger',            '✅ Live',        ''),
        ('AI Analyse (Step 2)',       '✅ Funktioniert','Thread-aware'),
        ('Paths vollst./unvollst.',   '✅ Konfiguriert',''),
        ('Kalkulation (Step 6)',      '✅ Deterministisch','Fixpreise'),
        ('Offerte (Step 7)',          '✅ Verknüpft',  'Mit Kalkulations-Preisen'),
        ('Human in the Loop (Step 9)','✅ Aktiv',      'Approval-Mail funktioniert'),
        ('Path nach HitL (11/13)',    '✅ Konfiguriert','approved / rejected'),
        ('Rückfrage-Mail (Step 16)',  '⚠️ Prüfen',     'To-Feld auf kunde_email-Chip'),
    ]
    alt = False
    for comp, status, detail in rows:
        row = table.add_row().cells
        row[0].text = comp
        row[1].text = status
        row[2].text = detail
        fill = 'F5F7FA' if alt else 'FFFFFF'
        for cell in row:
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill)
            cell._tc.get_or_add_tcPr().append(shd)
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10)
        alt = not alt

    for col_idx, width in enumerate([Cm(6), Cm(3.5), Cm(6)]):
        for row in table.rows:
            row.cells[col_idx].width = width

def add_learnings_table(doc):
    rows_data = [
        ('01', 'Fixwerte statt Ranges',      'AI-Kalkulationen brauchen exakte Preistabellen – keine Spielräume.'),
        ('02', 'Thread-Kontext ist alles',   'Der gesamte E-Mail-Verlauf muss ausgewertet werden, nicht nur die letzte Mail.'),
        ('03', 'Flache JSON-Struktur',       'Zapier braucht Top-Level-Felder – verschachtelte Objekte können nicht direkt gemappt werden.'),
        ('04', 'Chips ≠ Plaintext',          'In Zapier niemals {{...}} manuell tippen – immer den Chip-Picker verwenden.'),
        ('05', 'Input Fields explizit',      'Jeder AI-Step braucht jeden einzelnen Input-Wert als gemappten Chip.'),
        ('06', 'HitL nur live testbar',      'Approve/Reject funktioniert nur bei veröffentlichtem Zap mit echtem Trigger.'),
    ]
    table = doc.add_table(rows=0, cols=3)
    table.style = 'Table Grid'
    for num, title, desc in rows_data:
        row = table.add_row().cells
        p0 = row[0].paragraphs[0]
        r = p0.add_run(num)
        set_run_font(r, size=14, bold=True, color=ORANGE)

        p1 = row[1].paragraphs[0]
        r = p1.add_run(title)
        set_run_font(r, size=10, bold=True, color=DARK_BLUE)

        p2 = row[2].paragraphs[0]
        r = p2.add_run(desc)
        set_run_font(r, size=10, color=RGBColor(0x33, 0x33, 0x33))

        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F5F7FA')
        for cell in row:
            cell._tc.get_or_add_tcPr().append(copy.deepcopy(shd))

    for col_idx, width in enumerate([Cm(1.5), Cm(4.5), Cm(9.5)]):
        for row in table.rows:
            row.cells[col_idx].width = width

# ══════════════════════════════════════════════════════════════════
# COVER / TITLE
# ══════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(40)
p.paragraph_format.space_after  = Pt(6)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Projektjournal')
set_run_font(run, size=28, bold=True, color=DARK_BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
run = p.add_run('KI-gestützte Auftragsannahme')
set_run_font(run, size=20, bold=True, color=ORANGE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
run = p.add_run('Schellenberg Druck AG, Pfäffikon ZH')
set_run_font(run, size=13, italic=True, color=GREY)

add_divider(doc)

# ══════════════════════════════════════════════════════════════════
# AUSGANGSLAGE
# ══════════════════════════════════════════════════════════════════
add_heading(doc, 'Ausgangslage', 1)
add_body(doc, 'Die Schellenberg Druck AG erhält täglich Anfragen per E-Mail. Bisher musste jede Anfrage manuell gelesen, bewertet, kalkuliert und beantwortet werden – ein zeitaufwändiger, fehleranfälliger Prozess.')
add_body(doc, 'Das Ziel: Ein vollautomatisierter, KI-gestützter Workflow, der eingehende Kundenanfragen analysiert, Rückfragen stellt, Preise berechnet und fertige Angebote zur Freigabe vorlegt – ohne manuelle Eingriffe bis zum Freigabemoment.')
add_divider(doc)

# ══════════════════════════════════════════════════════════════════
# DIE REISE
# ══════════════════════════════════════════════════════════════════
add_heading(doc, 'Die Reise', 1)

# Phase 1
add_heading(doc, 'Phase 1 – Konzept und Grundstruktur', 2)
add_body(doc, 'Entscheidung: Zapier als Automatisierungsplattform, Gmail als Trigger, Claude (Anthropic) als AI-Engine.')
add_body(doc, 'Der Workflow wurde von Grund auf neu konzipiert:')
add_code(doc, 'Gmail-Eingang → AI Analyse → Vollständig?\n  → Nein: Automatische Rückfrage an Kunden\n  → Ja:   Kalkulation → Offerte → Human Approval → Angebot versenden')
add_body(doc, 'Definierte Pflichtfelder:')
add_bullet(doc, 'Name, Firma, Adresse, Telefon, E-Mail', 'Kundendaten:')
add_bullet(doc, 'Produkt, Auflage, Format, Farbigkeit, Material, Veredelung, Liefertermin, Druckdaten', 'Auftragsdaten:')

# Phase 2
add_heading(doc, 'Phase 2 – AI-Analyse-Prompt (Step 2)', 2)

add_body(doc, 'Problem 1: Redundante Rückfragen')
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Problem: ')
set_run_font(r, bold=True, color=RED_SOFT)
r2 = p.add_run('Kunde schreibt «4/0 farbig» – AI fragt trotzdem «Einseitig oder beidseitig?»')
set_run_font(r2, size=11)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Lösung: ')
set_run_font(r, bold=True, color=GREEN)
r2 = p.add_run('Interpretationsregeln eingebaut: 4/0 = einseitig (klar), 4/4 = beidseitig, «keine Veredelung» direkt eintragen.')
set_run_font(r2, size=11)

doc.add_paragraph()
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Problem: ')
set_run_font(r, bold=True, color=RED_SOFT)
r2 = p.add_run('Bei Follow-up-Mails las die KI nur die neueste Nachricht.')
set_run_font(r2, size=11)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Lösung: ')
set_run_font(r, bold=True, color=GREEN)
r2 = p.add_run('Prompt auf vollständiges Thread-Reading umgestellt. Bei Widersprüchen gilt die neuere Angabe.')
set_run_font(r2, size=11)

doc.add_paragraph()
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Problem: ')
set_run_font(r, bold=True, color=RED_SOFT)
r2 = p.add_run('E-Mail-Adresse fehlte im Output.')
set_run_font(r2, size=11)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Lösung: ')
set_run_font(r, bold=True, color=GREEN)
r2 = p.add_run('E-Mail wird automatisch aus dem Absender-Feld (absender_email) übernommen.')
set_run_font(r2, size=11)

doc.add_paragraph()
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Neues Feature: ')
set_run_font(r, bold=True, color=ORANGE)
r2 = p.add_run('Bei vollständigem Auftrag automatische Eingangsbestätigung («Wir melden uns in 1–2 Werktagen»).')
set_run_font(r2, size=11)

# Phase 3
add_heading(doc, 'Phase 3 – Zapier-Struktur und Pfade', 2)
add_bullet(doc, 'status = unvollstaendig → Rückfrage-Mail an Kunden', 'Pfad A:')
add_bullet(doc, 'status = vollstaendig → Kalkulation → Offerte → Freigabe', 'Pfad B:')
doc.add_paragraph()
add_body(doc, 'Erkenntnis: Zapier-Paths haben keine Formatter-Steps. Datentransformation muss vor dem Path oder im AI-Prompt selbst erfolgen.')
add_body(doc, 'Zapier-Kompatibilität: Das JSON-Output enthält flache Top-Level-Felder (kunde_name, kunde_strasse, etc.) zusätzlich zum verschachtelten kunde-Objekt.')

# Phase 4
add_heading(doc, 'Phase 4 – Kalkulation (Step 6)', 2)
add_body(doc, 'Erste Version mit Preisspannen führte zu zufälligen Selbstkosten pro Run. Lösung: Vollständige Preistabelle mit exakten Fixwerten.')
add_code(doc,
'Gesamtbogen  =  AUFRUNDEN((Auflage × 1.08) / 50) × 50\n'
'Selbstkosten =  Papier + Druck + Veredelung + Layout + Express\n'
'VP netto     =  Selbstkosten × 1.7241  →  aufrunden auf CHF 5.–  (min. CHF 95.–)\n'
'MwSt.        =  VP netto × 0.081')
add_body(doc, 'Handelsmarge: 42% auf Nettoverkaufspreis = 72.41% Aufschlag auf Selbstkosten.')
add_body(doc, 'Output: Detaillierter JSON-Block mit Kalkulation, beiden Margen-Kennzahlen, MwSt. 8.1%, Preis pro Stück und Optimierungshinweisen.')

# Phase 5
add_heading(doc, 'Phase 5 – Offerte (Step 7)', 2)
add_body(doc, 'Claude erstellt ein rechtssicheres Schweizer Geschäftsdokument mit Briefkopf, Empfänger, Leistungsbeschreibung, Preisen, Liefer-/Zahlungsbedingungen und Unterschrift Edgar Glinz.')
add_body(doc, 'Layout-Regeln: Kein Markdown, kein **, kein ---, keine Tabellen, max. 2 Seiten, max. 1 Leerzeile zwischen Abschnitten.')
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Problem: ')
set_run_font(r, bold=True, color=RED_SOFT)
r2 = p.add_run('Step 7 hatte keinen Zugriff auf die Preise aus Step 6.')
set_run_font(r2, size=11)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Lösung: ')
set_run_font(r, bold=True, color=GREEN)
r2 = p.add_run('Input Fields in Step 7 um 4 Felder ergänzt: verkaufspreis_netto, mwst_betrag, verkaufspreis_brutto, preis_pro_stueck_netto.')
set_run_font(r2, size=11)

# Phase 6
add_heading(doc, 'Phase 6 – Human in the Loop (Step 9)', 2)
add_body(doc, 'Edgar Glinz erhält eine Approval-E-Mail mit dem fertigen Angebot. Approve → Angebot an Kunden. Reject → Ablehnungs-Mail mit Begründung.')
add_bullet(doc, '«Content to review» war auf Kunden-Mail (Step 1) statt Step 7 gemappt → falscher Inhalt')
add_bullet(doc, 'Approve-Link leitet im Test auf Zap-Editor → normal, funktioniert nur live')
add_bullet(doc, 'Path-Bedingungen: Status = «approved» / «rejected» (Grossbuchstabe, Exactly matches)')
add_bullet(doc, 'Zapbot-Benachrichtigung: Zapier-System-Mail → deaktivierbar unter zapier.com/app/settings/notifications')

# Phase 7
add_heading(doc, 'Phase 7 – End-to-End-Test', 2)
add_body(doc, 'Erster Live-Test mit realer Anfrage (Patrick Glinz, Event-Werbematerial: Flyer, Plakate, Pappbecher). AI las den gesamten E-Mail-Thread korrekt, erkannte korrigierten Liefertermin.')
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Fehler: ')
set_run_font(r, bold=True, color=RED_SOFT)
r2 = p.add_run('Step 16 «To»-Feld hatte {{Trigger – From Email}} als Plaintext → «Recipient address required».')
set_run_font(r2, size=11)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
r = p.add_run('Fix: ')
set_run_font(r, bold=True, color=GREEN)
r2 = p.add_run('Chip auf kunde_email aus Step 2 umgestellt.')
set_run_font(r2, size=11)

add_divider(doc)

# ══════════════════════════════════════════════════════════════════
# AKTUELLER STAND
# ══════════════════════════════════════════════════════════════════
add_heading(doc, 'Aktueller Stand', 1)
add_status_table(doc)
doc.add_paragraph()
add_divider(doc)

# ══════════════════════════════════════════════════════════════════
# KEY LEARNINGS
# ══════════════════════════════════════════════════════════════════
add_heading(doc, 'Key Learnings', 1)
add_learnings_table(doc)
doc.add_paragraph()
add_divider(doc)

# ══════════════════════════════════════════════════════════════════
# NÄCHSTE SCHRITTE
# ══════════════════════════════════════════════════════════════════
add_heading(doc, 'Nächste Schritte', 1)
next_steps = [
    'To-Feld Step 16: Chip auf kunde_email umstellen',
    'Vollständiger Live-Test: Auftrag annehmen + Angebot-Versand prüfen',
    'Google Docs: PDF-Anhang für das Angebot',
    'Multi-Produkt-Aufträge validieren (Flyer + Plakat + Sonderprodukt)',
    'Sonderprodukte ergänzen (Pappbecher, Textilien etc.)',
]
for step in next_steps:
    add_bullet(doc, step)

add_divider(doc)

# Footer note
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Schellenberg Druck AG · Schützenhausstrasse 5 · 8330 Pfäffikon · schellenbergdruck.ch · Juni 2026')
set_run_font(run, size=9, italic=True, color=GREY)

# ── Save ───────────────────────────────────────────────────────
out = '/home/user/zapier-json-schema/Journal_KI_Auftragsannahme.docx'
doc.save(out)
print(f'Saved: {out}')
