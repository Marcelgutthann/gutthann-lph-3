---
name: lph3
description: Erstellt einen vollstaendigen LPH 3 Entwurfsplanungsbericht (HOAI Paragraph 33, Gebaeude und Innenraeume) fuer ein beliebiges Architekturprojekt. Durchsucht alle Dateien im Projektordner, extrahiert Daten und befuellt das standardisierte HTML-Template. Baut auf der LPH 2 Dokumentation auf.
user-invocable: true
---

# LPH 3 Entwurfsplanungsbericht — Automatische Generierung

## OUTPUT-KONVENTION (verbindlich)

**WICHTIGSTE REGEL: Es wird ein EINZIGER Ordner namens `Claude/` im Projekt-Root erzeugt. ALLE Outputs landen dort drin. NIRGENDS sonst im Projekt.**

```
<Projekt-Root>/                           <- bleibt unveraendert
├── (die existierenden Projektordner)       (welche Struktur auch immer)
│
└── Claude/                               <- EINZIGER Ordner den wir anlegen
    ├── Projekt_Historie.md                 <- kumulative Wissensbasis
    ├── LPH_1/LPH1_Report_YYYY-MM-DD.html
    ├── LPH_2/LPH2_Report_YYYY-MM-DD.html
    ├── ... bis LPH_8/
    ├── Dashboard/Projekt_Dashboard_YYYY-MM-DD.html
    └── Fachplaner-Analyse/Fachplaner_Analyse_YYYY-MM-DD.md
```

**Dein konkreter Output-Pfad fuer diesen Skill:**
```
<Projekt-Root>/Claude/LPH_3/LPH3_Report_YYYY-MM-DD.html
```

**Pflichten beim Schreiben:**
1. Wenn `<Projekt-Root>/Claude/` nicht existiert: anlegen (mkdir)
2. Wenn `<Projekt-Root>/Claude/LPH_3` nicht existiert: anlegen
3. Datums-Format im Dateinamen: `YYYY-MM-DD` (z.B. 2026-05-27)
4. Bei jedem Aufruf: NEUE Datei mit aktuellem Datum (keine Ueberschreibung)

**NIEMALS schreibst du:**
- In andere Ordner des Projekts (00 Aktennotiz/, 11 Doku/, etc.) - keine Verschmutzung der Projekt-Struktur
- In den Projekt-Root direkt - nur in `Claude/` Unterstruktur
- In `~/.claude/plugins/` - das ist Read-only Plugin-Code
- In versteckte Ordner (.claude/) - nutze ausschliesslich `Claude/` (sichtbar)

---

## WISSENSBASIS NUTZEN

Falls `<Projekt-Root>/Claude/Projekt_Historie.md` existiert:
- ZUERST diese Datei lesen - sie enthaelt chronologisch alle E-Mails, Aktennotizen, Beschluesse
- Spart Zeit + Tokens (statt das ganze Projekt jedes Mal neu zu scannen)

Falls die Datei NICHT existiert:
- Empfehle dem User: "Tipp: Fuer aktuelle Daten zuerst /projekt-historie ausfuehren"
- Trotzdem das Projekt direkt scannen (Fallback)

---


## Template
Suche das Template relativ zum aktuellen Projektordner — KEIN hardcodierter Pfad:
```powershell
# Template finden:
Get-ChildItem -Path "." -Recurse -Filter "LPH3_Template.html" | Select-Object -First 1 -ExpandProperty FullName
```
Das Template liegt üblicherweise in `Claude_DOKU\LPH3_Template.html` (aus dem _Claude_DOKU_Starter).
Kopiere es als `Claude_DOKU\LPH3_Report.html` und befuelle alle `{{PLATZHALTER}}`.
Falls nicht gefunden: Benutzer informieren dass das Template fehlt.

## Ablauf — 6 Schritte

### Schritt 1: LPH 2 Dokumentation finden (PFLICHT)
Durchsuche den Projektordner nach der bestehenden LPH 2 Dokumentation:
- Suche nach `*LPH2*`, `*LPH_2*`, `*Vorentwurf*`, `*Doku*LPH*2*`
- Lies die LPH 2 HTML-Datei oder PDF vollstaendig
- Extrahiere daraus: Projektziele, Kostenschaetzung, Raumprogramm, Beteiligte, Terminplan, Zielkonflikte
- Diese Daten sind die GRUNDLAGE fuer den LPH 3 Bericht (Fortschreibung!)

### Schritt 2: Projektordner durchsuchen
Durchsuche den GESAMTEN Projektordner jede Datei Seite fuer Seite rekursiv. DU DARFST NICHTS UEBERSPRINGEN:
- **PDFs**: Lies JEDE Seite jeder PDF (Glob `**/*.pdf`, dann Read)
- **DOCX**: Lies per PowerShell Word COM (`New-Object -ComObject Word.Application`)
- **XLSX**: Lies per PowerShell Excel COM (`New-Object -ComObject Excel.Application`)
- **Bilder**: Finde Logo (GHIW/Gutthann/Logo im Namen) und Titelbild AKTUELLE VERSION (Perspektive/Visu/Lageplan)
- **Jour Fixe PDFs**: Suche nach `*JourFixe*`, `*Jour*Fixe*`, `*PS-JF*`, `*FPJF*`
- **Beteiligtenliste**: Suche nach `*Beteiligte*` oder `*Projektbeteiligte*`
- **LPH 2 Dokumentation**: Bereits erledigte Dokumentation ueber die vorangegangene Leistungsphase (PFLICHT-Quelle!)
- **Entwurfsplaene**: Suche nach Dateien im Ordner `*Entwurf*` oder `aktueller Stand`
- **Kostenberechnung**: Suche nach `*Kostenberechnung*`, `*KBe*`, `*DIN 276*` in Ordner `09 Kosten`

WICHTIG: DOCX und XLSX koennen NICHT mit dem Read-Tool gelesen werden. Verwende IMMER PowerShell COM.

### Schritt 3: Daten extrahieren
Extrahiere aus allen gelesenen Dateien:

**Vorgaben aus LPH 2 (Bruecke zur Vorplanung):**
- Was wurde in LPH 2 festgelegt? (Varianten, Entscheidungen, offene Punkte)
- Kostenschaetzung aus LPH 2 als Vergleichsbasis
- Raumprogramm Soll aus LPH 2
- Terminplan aus LPH 2

**Projektdaten LPH 3:**
- **Beteiligte**: Namen, Firmen, Kontakte (NUR aus Beteiligtenliste, NICHTS erfinden)
- **Flaechen**: BGF, BRI, NUF je Bauteil
- **Raumprogramm**: Soll/Ist-Vergleich je Raum (detaillierter als LPH 2)
- **Kosten**: Kostenberechnung DIN 276, KGR 100-700, Vergleich mit Kostenschaetzung LPH 2
- **Termine**: Fortgeschriebener Terminplan, Meilensteine, Bauabschnitte
- **Brandschutz**: Konzept, Nachweis, offene Punkte
- **Barrierefreiheit**: Anforderungen, Massnahmen, Abstimmungen
- **TGA**: Lueftung, Heizung, Elektro, Bauphysik — Verweis auf Fachplaner-Erlaeuterungsberichte
- **B-Plan**: Planungsrechtliche Zulaessigkeit, Befreiungen
- **Objektbeschreibung**: KGR 100-700 detailliert (Konstruktion, Materialien, Ausfuehrung)
- **Beschluesse aus Jour Fixe**: Entscheidungen, offene Punkte, Action Items

### Schritt 4: Template kopieren und befuellen
1. Kopiere `LPH3_Template.html` in den `Claude_DOKU` Ordner als `LPH3_Report.html`
2. Ersetze ALLE `{{PLATZHALTER}}` mit den extrahierten Daten
3. Erstelle HTML-Tabellen fuer Tabellen-Platzhalter (Raumprogramm, Kosten, Objektbeschreibung etc.)
4. Befuelle die Sidebar-Dateiliste mit allen tatsaechlich gelesenen Dateien
5. Objektbeschreibung KGR 100-700: Fuer JEDE Kostengruppe detaillierte Baubeschreibung

### Schritt 5: Planungsfortschritt analysieren
Analysiere den Fortschritt von LPH 2 zu LPH 3:
- Was hat sich seit dem Vorentwurf geaendert?
- Welche Variante wurde gewaehlt und warum?
- Kostenentwicklung: Kostenschaetzung (LPH 2) vs. Kostenberechnung (LPH 3)
- Raumprogramm: Aenderungen seit LPH 2
- Neue Erkenntnisse (Brandschutz, Baugrund, TGA etc.)
- Erstelle Vergleichstabelle: Kostenschaetzung vs. Kostenberechnung je KGR

### Schritt 6: Qualitaetspruefung
- Keine `{{PLATZHALTER}}` mehr im Output — pruefe mit PowerShell:
  `Select-String -Path "LPH3_Report.html" -Pattern "\{\{" | Measure-Object`
- Alle Zahlen gegen Quelldokumente pruefen
- Schreibstil: Sachlich, technisch, Passiv, unpersoenlich ("ist einzuhalten", "wird geprueft")
- Max. 5-8 Saetze pro Absatz
- Beteiligte nur indirekt ("seitens AG", "durch Fachplaner")
- Objektbeschreibung muss ALLE KGR 100-700 abdecken

## Design-Regeln
- Farbe: SCHWARZ `#1a1a1a` — kein Gruen, kein Teal, kein Blau
- Alle Felder `contenteditable` fuer nachtraegliche Bearbeitung
- PDF-Druck: A4 Querformat, randlos, keine Schatten, keine Checklisten
- Objektbeschreibung KGR als dreispaltiges Layout (wie Beispiel-PDFs)

## Platzhalter-Referenz

| Platzhalter | Beschreibung |
|---|---|
| `{{PROJEKT_TITEL}}` | z.B. "NEUBAU GRUNDSCHULE MIT OGTS" |
| `{{PROJEKT_NR}}` | z.B. "(Projekt-Nr.)-24" |
| `{{PROJEKT_ORT}}` | z.B. "{ORTSNAME}" |
| `{{DATUM}}` | Aktuelles Datum TT.MM.JJJJ |
| `{{TITELBILD_PFAD}}` | Relativer Pfad zum Titelbild |
| `{{AG_NAME}}` | Name des Auftraggebers |
| `{{3_1_1_ALLGEMEINES}}` | Fliesstext Allgemeines zur Entwurfsplanung |
| `{{3_1_2_STAEDTEBAU}}` | Fliesstext Staedtebau / Erschliessung / Architektur |
| `{{3_1_3_PLANUNGSRECHT}}` | Fliesstext Planungsrechtliche Zulaessigkeit |
| `{{3_1_4_RAUMBEDARF}}` | Fliesstext Erfuellung des Raumbedarfs |
| `{{3_1_5_OEFFENTL_RECHT}}` | Fliesstext Oeffentlich-rechtliche Anforderungen |
| `{{3_1_6_ERWEITERUNG}}` | Fliesstext Erweiterungsmoeglichkeit |
| `{{3_1_7_KOSTENSICHERHEIT}}` | Fliesstext Kostensicherheit / -risiko |
| `{{3_1_8_BARRIEREFREIHEIT}}` | Fliesstext Barrierefreiheit |
| `{{3_1_9_NUTZERPROFIL}}` | Fliesstext Nutzerprofil |
| `{{3_2_ARBEITSERGEBNISSE}}` | Fliesstext Bereitstellen der Arbeitsergebnisse |
| `{{3_2_PLANVERZEICHNIS_TABELLE}}` | HTML Tabelle Planverzeichnis |
| `{{3_3_KGR100}}` bis `{{3_3_KGR700}}` | Objektbeschreibung je KGR als HTML |
| `{{3_4_GENEHMIGUNG}}` | Fliesstext Verhandlungen Genehmigungsfaehigkeit |
| `{{3_5_KOSTEN_TEXT}}` | Einleitungstext Kostenberechnung |
| `{{3_5_KOSTEN_TABELLE}}` | HTML Kostentabelle DIN 276 |
| `{{3_5_KOSTEN_VERGLEICH}}` | HTML Vergleichstabelle KS vs. KB |
| `{{3_6_TERMINPLAN_TEXT}}` | Einleitungstext Terminplan |
| `{{3_6_TERMINPLAN_TABELLE}}` | HTML Terminplan-Tabelle |
| `{{3_7_ZUSAMMENFASSUNG}}` | Fliesstext Zusammenfassung |
| `{{3_8_BESONDERE_LEISTUNGEN}}` | Liste besondere Leistungen |
| `{{3_1_BETEILIGTE_TABELLE}}` | HTML Tabelle fachlich Beteiligte |
| `{{ABSCHLUSS_TEXT}}` | Abschlusstext LPH 3 |
| `{{DATEILISTE}}` | `<li>` Elemente fuer Sidebar |
