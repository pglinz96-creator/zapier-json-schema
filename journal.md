# Projektjournal – KI-gestützte Auftragsannahme
## Schellenberg Druck AG, Pfäffikon ZH

---

## Ausgangslage

Die Schellenberg Druck AG erhält täglich Anfragen per E-Mail. Bisher musste jede Anfrage manuell gelesen, bewertet, kalkuliert und beantwortet werden – ein zeitaufwändiger, fehleranfälliger Prozess.

Das Ziel: Ein vollautomatisierter, KI-gestützter Workflow, der eingehende Kundenanfragen analysiert, Rückfragen stellt, Preise berechnet und fertige Angebote zur Freigabe vorlegt – ohne manuelle Eingriffe bis zum Freigabemoment.

---

## Die Reise

### Phase 1 – Konzept und Grundstruktur

**Entscheidung:** Zapier als Automatisierungsplattform, Gmail als Trigger, Claude (Anthropic) als AI-Engine.

Der Workflow wurde von Grund auf neu konzipiert:

```
Gmail-Eingang → AI Analyse → Vollständig?
  → Nein: Automatische Rückfrage an Kunden
  → Ja:   Kalkulation → Offerte → Human Approval → Angebot versenden
```

Die erste grosse Herausforderung: Was ist ein "vollständiger" Auftrag? Welche Pflichtangaben braucht eine Druckerei wirklich?

**Definierte Pflichtfelder:**
- Kundendaten: Name, Firma, Adresse, Telefon, E-Mail
- Auftragsdaten: Produkt, Auflage, Format, Farbigkeit, Material, Veredelung, Liefertermin, Druckdaten vorhanden?

---

### Phase 2 – Der AI-Analyse-Prompt (Step 2)

Der erste Prompt war zu einfach. Probleme tauchten auf:

**Problem:** Die KI stellte Rückfragen für Angaben, die längst beantwortet waren.
Beispiel: Kunde schreibt "4/0 farbig" – AI fragt trotzdem "Einseitig oder beidseitig?"

**Lösung:** Interpretationsregeln eingebaut:
- `4/0` = vierfarbig einseitig → klar, nicht nochmal fragen
- `4/4` = beidseitig → klar
- "keine Veredelung" / kein Hinweis → als "keine" eintragen

**Problem:** Bei Follow-up-Mails las die KI nur die neueste Nachricht.
**Lösung:** Prompt auf vollständiges Thread-Reading umgestellt – alle Mails im Verlauf werden ausgewertet, bei Widersprüchen gilt die neuere Angabe.

**Problem:** E-Mail-Adresse fehlte im Output.
**Lösung:** E-Mail wird automatisch aus dem Absender-Feld (`absender_email`) übernommen – niemals manuell gesucht oder nachgefragt.

**Problem:** Folge-E-Mails wirkten zu förmlich (lange Dankesreden).
**Lösung:** Tonregel eingebaut: Erste Nachricht = freundlich mit Einleitung. Folgeantworten = kurz, direkt, sachlich.

**Neues Feature:** Bei vollständigem Auftrag wird automatisch eine Eingangsbestätigung generiert ("Ihr Auftrag ist vollständig eingegangen, wir melden uns in 1–2 Werktagen").

---

### Phase 3 – Zapier-Struktur und Pfade

**Step 3 – Paths (Verzweigung):**
- Pfad A: `status = unvollstaendig` → Rückfrage-Mail an Kunden
- Pfad B: `status = vollstaendig` → Kalkulation → Offerte → Freigabe

**Erkenntnis:** Zapier-Paths haben keine Formatter-Steps. Datentransformation muss vor dem Path passieren oder im AI-Prompt selbst erfolgen.

**Flache Felder für Zapier:** Das JSON-Output der AI enthält sowohl ein verschachteltes `kunde`-Objekt als auch flache Top-Level-Felder (`kunde_name`, `kunde_strasse`, `kunde_plz`, `kunde_ort`), weil Zapier nur flache Felder direkt mappen kann.

---

### Phase 4 – Kalkulation (Step 6)

**Erste Version:** Ranges ("CHF 18–25 pro 1000 Bogen") → AI wählte zufällige Werte → inkonsistente Preise.

**Problem:** Bei jedem Testlauf andere Selbstkosten, obwohl Marge (42%) immer stimmte.

**Lösung:** Vollständige Preistabelle mit exakten Fixwerten:
- Papier: CHF 18.–/1000 Bogen (90g Offset) bis CHF 95.–/1000 Bogen (Folie)
- Rüstkosten: fix CHF 55.–
- Druckkosten: fix CHF 0.048/Bogen (≤1000), CHF 0.034 (1001–5000), CHF 0.024 (>5000)
- Beidseitig: Faktor ×1.75
- Veredelung: feste CHF-Beträge pro Bogen

**Kalkulations-Formel:**
```
Gesamtbogen = AUFRUNDEN((Auflage × 1.08) / 50) × 50
Selbstkosten = Papier + Druck + Veredelung + Layout + Express
VP netto = SK × 1.7241  →  aufrunden auf CHF 5.–  (min. CHF 95.–)
```

**Handelsmarge:** 42% auf Nettoverkaufspreis = 72.41% Aufschlag auf Selbstkosten.

**Output:** Detaillierter JSON-Block mit Kalkulation, beiden Margen-Kennzahlen, MwSt. 8.1%, Preis pro Stück und Optimierungshinweisen.

---

### Phase 5 – Offerte (Step 7)

Der Offerte-Prompt erstellt ein rechtssicheres Schweizer Geschäftsdokument.

**Pflichtbestandteile:**
- Briefkopf Schellenberg Druck AG
- Empfänger rechtsbündig
- Leistungsbeschreibung mit Preisen (netto / MwSt. / brutto)
- Liefer- und Zahlungsbedingungen
- Rechtliche Hinweise (kompakt, unauffällig)
- Unterschrift Edgar Glinz

**Layout-Regeln:** Kein Markdown, kein `**`, kein `---`, keine Tabellen, max. 2 Seiten, max. 1 Leerzeile zwischen Abschnitten.

**Herausforderung:** Step 7 hatte keinen Zugriff auf die Preise aus Step 6.
**Lösung:** Input Fields in Step 7 um 4 Felder aus Step 6 ergänzt: `verkaufspreis_netto`, `mwst_betrag`, `verkaufspreis_brutto`, `preis_pro_stueck_netto`.

---

### Phase 6 – Human in the Loop (Step 9)

Edgar Glinz erhält eine E-Mail mit dem fertigen Angebot zur Freigabe.

**Konfiguration:**
- Content to review: Step 7 Output (Offerte-Text)
- Approver: p.glinz96@gmail.com
- Nach Approve → Angebot per Mail an Kunden
- Nach Reject → Ablehnungs-Mail mit Begründung an Kunden

**Stolpersteine:**
- "Content to review" war auf Kunden-Mail (Step 1) gemappt statt auf Step 7 → Falscher Inhalt in der Approval-Mail
- Während Zapier-Tests zeigt Approve-Link auf den Zap-Editor → normal, funktioniert nur bei veröffentlichtem, live-getriggertem Zap
- Path-Bedingungen: `Approved` / `Rejected` exakt mit Grossbuchstabe, Operator "Exactly matches"

---

### Phase 7 – End-to-End-Test

Erster erfolgreicher Live-Test mit realer Kundenanfrage (Patrick Glinz, Event-Werbematerial):
- Anfrage enthielt mehrere Produkte (Flyer, Plakate, Pappbecher), Liefertermin war korrigiert
- AI las den gesamten E-Mail-Thread korrekt aus
- Unvollständiger Auftrag → Rückfrage wurde gesendet
- Fehler: Step 16 "To"-Feld hatte `{{Trigger – From Email}}` als Plaintext statt Chip → "Recipient address required"
- Fix: Chip auf `kunde_email` aus Step 2 umgestellt

---

## Aktueller Stand

| Komponente | Status |
|---|---|
| Gmail-Trigger | ✅ Live |
| Step 2 AI Analyse | ✅ Funktioniert |
| Step 3 Paths | ✅ Konfiguriert |
| Step 6 Kalkulation | ✅ Fixpreise, deterministisch |
| Step 7 Offerte | ✅ Mit Kalkulations-Preisen verknüpft |
| Step 9 Human in the Loop | ✅ Approval-Mail funktioniert |
| Step 11/13 Path nach HitL | ⚠️ Condition-Werte prüfen (Approved/Rejected) |
| Step 16 Rückfrage-Mail | ⚠️ To-Feld auf kunde_email-Chip umstellen |

---

## Gelernte Lektionen

1. **Fixwerte statt Ranges:** AI-Kalkulationen müssen mit exakten Werten arbeiten, keine Interpretationsspielräume.
2. **Thread lesen:** Der Prompt muss explizit angewiesen werden, den gesamten E-Mail-Verlauf auszuwerten.
3. **Flache JSON-Felder:** Zapier braucht Top-Level-Felder – keine verschachtelten Objekte für Field-Mapping.
4. **Chips statt Text:** In Zapier niemals `{{...}}` manuell eintippen – immer den Chip-Picker verwenden.
5. **Input Fields:** Jeder AI-Schritt braucht explizit konfigurierte Input Fields mit gemappten Chips.
6. **HitL funktioniert nur live:** Tests im Zapier-Editor simulieren den Approve-Link nicht korrekt.
7. **Tone of Voice:** Rückfrage-Prompts brauchen explizite Ton-Regeln für Erst- und Folgeantworten.

---

## Nächste Schritte

- [ ] Step 16 "To"-Chip auf `kunde_email` umstellen
- [ ] Step 11/13 Path-Conditions mit `Approved`/`Rejected` (Grossschreibung) prüfen
- [ ] Vollständiger Live-Test mit einem vollständigen Auftrag
- [ ] HitL Approve → Angebot an Kunden testen
- [ ] Optional: Google Docs als PDF-Anhang für das Angebot

---

*Projekt gestartet: Juni 2026*
*Schellenberg Druck AG, Schützenhausstrasse 5, 8330 Pfäffikon*
