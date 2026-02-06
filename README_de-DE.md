[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

Eine Software zur Konvertierung von Dokumenten- und Diagrammformaten - Unterstützt die bidirektionale Konvertierung von Word/Markdown/Excel. Läuft vollständig lokal und gewährleistet Datensicherheit und Zuverlässigkeit.

## 📖 Projekthintergrund

Diese Software wurde ursprünglich für die tägliche Arbeit der Druckerei entwickelt, um folgende Probleme zu lösen:
- Die von verschiedenen Abteilungen gesendeten Dokumentformate sind chaotisch und müssen in standardisierte Formate organisiert werden.
- Es gibt viele Arten von Dokumenten, jede mit unterschiedlichen festen Formatanforderungen.
- Muss offline laufen und sich an Intranet-Umgebungen und ältere Geräte anpassen.

**Designphilosophie**: Diese Software ist als leichtes, narrensicheres Werkzeug positioniert. Obwohl sie in Bezug auf Professionalität und funktionale Vollständigkeit nicht mit professionellen Werkzeugen wie LaTeX oder Pandoc verglichen werden kann, zeichnet sie sich durch null Lernkosten und sofortige Einsatzbereitschaft aus, was sie für tägliche Büroszenarien geeignet macht, in denen die Formatanforderungen nicht extrem streng sind.

## ✨ Kernfunktionen

- **📄 Dokumentformatkonvertierung** - Bidirektionale Word ↔ Markdown Konvertierung. Unterstützt mathematische Formelkonvertierung und bidirektionale Trennzeichenkonvertierung (Markdowns drei Arten von Trennlinien vs. Words Seitenumbrüche, Abschnittswechsel und horizontale Linien). Unterstützt Formate wie DOCX/DOC/WPS/RTF/ODT.
- **📊 Tabellenformatkonvertierung** - Bidirektionale Excel ↔ Markdown Konvertierung. Unterstützt XLSX/XLS/ET/ODS/CSV Formate. Enthält Tabellenzusammenfassungswerkzeuge.
- **📑 PDF und Layoutdateien** - PDF/XPS/OFD zu Markdown oder DOCX Konvertierung. Unterstützt PDF-Zusammenführung, -Teilung und andere Operationen.
- **🖼️ Bildverarbeitung** - Unterstützt bidirektionale Konvertierung und Komprimierung von JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC Formaten.
- **🔍 OCR-Texterkennung** - Integriertes RapidOCR zum Extrahieren von Text aus Bildern und PDFs.
- **✏️ Textkorrektur** - Überprüft auf Tippfehler, Zeichensetzung, Symbole und sensitive Wörter basierend auf benutzerdefinierten Wörterbüchern. Regeln können in der Einstellungsoberfläche bearbeitet werden.
- **📝 Vorlagensystem** - Flexibler Vorlagenmechanismus, der benutzerdefinierte Dokument- und Berichtsformate unterstützt.
- **💻 Dual-Modus-Betrieb** - Grafische Benutzeroberfläche (GUI) + Befehlszeilenschnittstelle (CLI).
- **🔒 Vollständig lokaler Betrieb** - Läuft offline und gewährleistet Datensicherheit mit integrierten Netzwerkisolationsmechanismen.
- **🔗 Einzelinstanzbetrieb** - Verwaltet automatisch Programminstanzen und unterstützt die Integration mit dem begleitenden Obsidian-Plugin.

## 📸 Screenshots

| Batch | Markdown |
| --- | --- |
| ![Batch-Ansicht](assets/screenshots/batch.png) | ![Markdown-Ansicht](assets/screenshots/markdown.png) |

| Dokument | Tabelle |
| --- | --- |
| ![Dokument-Ansicht](assets/screenshots/document.png) | ![Tabellen-Ansicht](assets/screenshots/spreadsheet.png) |

| Bild | Layout-Dateien |
| --- | --- |
| ![Bild-Ansicht](assets/screenshots/image.png) | ![Layout-Ansicht](assets/screenshots/layout.png) |

Changelog: siehe [doc/CHANGELOG.md](doc/CHANGELOG.md)

## 🚀 Schnellstart

### Programm starten

Doppelklicken Sie auf `DocWen.exe`, um die grafische Oberfläche zu starten.

### Schnellstartanleitung

1.  **Bereiten Sie eine Markdown-Datei vor**:

    ```markdown
    ---
    Titel: Testdokument
    ---
    
    ## Testtitel
    
    Dies ist der Testinhalt.
    ```

2.  **Drag & Drop Konvertierung**:
    - Starten Sie das Programm.
    - Ziehen Sie die `.md`-Datei in das Fenster.
    - Wählen Sie eine Vorlage.
    - Klicken Sie auf "In DOCX konvertieren".

3.  **Ergebnis erhalten**:
    - Ein standardisiertes Word-Dokument wird im selben Verzeichnis generiert.

**Tipp**: Sie können die Beispieldateien im Verzeichnis `samples/` verwenden, um die Funktionen der Software schnell auszuprobieren.

## 📝 Markdown-Syntaxkonventionen

### Überschriftenebenen-Zuordnung

Um es Kollegen ohne Hintergrundwissen leichter zu machen, entsprechen die Markdown-Überschriften in dieser Software **eins-zu-eins** den Word-Überschriften:
- Dokumenttitel (title) und Untertitel (subtitle) werden in den YAML-Metadaten platziert.
- Markdown `# Überschrift 1` entspricht Word "Überschrift 1".
- Markdown `## Überschrift 2` entspricht Word "Überschrift 2".
- Und so weiter, bis zu 9 Überschriftenebenen werden unterstützt.

**Tipp**: Wenn Sie es bevorzugen, Markdowns erste Überschriftenebene (`#`) als Dokumenttitel zu verwenden und ab der zweiten Ebene (`##`) mit Textüberschriften zu beginnen, können Sie den Stil "Überschrift 1" in der Word-Vorlage so anpassen, dass er wie ein Dokumenttitel aussieht (z.B. zentriert, fett, größere Schriftgröße), und in den Einstellungen ein Nummerierungsschema auswählen, das die Nummerierung der ersten Überschriftenebene überspringt. So erscheinen Ihre Überschriften der ersten Ebene als Dokumenttitel.

### Zeilenumbrüche und Absätze

**Grundregel**: Jede nicht leere Zeile wird standardmäßig als separater Absatz behandelt.

**Gemischte Absätze**: Wenn ein Untertitel mit dem Fließtext im selben Absatz gemischt werden muss, müssen folgende Bedingungen erfüllt sein:
1.  Der Untertitel endet mit einem Satzzeichen (unterstützt mehrsprachige Satzzeichen, einschließlich Punkte, Fragezeichen, Ausrufezeichen und andere gängige Schlusssatzzeichen).
2.  Der Fließtext befindet sich in der **unmittelbar nächsten Zeile** des Untertitels.
3.  Die Fließtextzeile darf kein spezielles Markdown-Element sein (wie Überschriften, Codeblöcke, Tabellen, Listen, Zitate, Formelblöcke, Trennzeichen usw.).

**Beispiel**:
```markdown
## I. Arbeitsanforderungen.
Dieses Treffen erfordert, dass alle Einheiten ernsthaft umsetzen...
```
Die obigen zwei Zeilen werden zu einem Absatz zusammengeführt, wobei "I. Arbeitsanforderungen." das Untertitelformat behält und "Dieses Treffen..." das Fließtextformat behält.

**Hinweis**:
- Zwischen Untertitel und Fließtext darf keine Leerzeile stehen; andernfalls werden sie als separate Absätze erkannt.
- Wenn der Untertitel nicht mit einem Satzzeichen endet und keine Leerzeile zum Fließtext hat, wird der Fließtext in die Überschriftenzeile mit angepasster Formatierung zusammengeführt.

### Bidirektionale Trennzeichenkonvertierung

Unterstützt die bidirektionale Konvertierung zwischen Markdown-Trennzeichen und Word-Seitenumbrüchen/Abschnittswechseln/horizontalen Linien:

-   **DOCX → MD**: Word-Seitenumbrüche, Abschnittswechsel und horizontale Linien werden automatisch in Markdown-Trennzeichen konvertiert.
-   **MD → DOCX**: Markdown `---`, `***`, `___` werden automatisch in entsprechende Word-Elemente konvertiert.
-   **Konfigurierbar**: Spezifische Zuordnungsbeziehungen können in der Einstellungsoberfläche angepasst werden.

### Bildeinbettung und Größe

Unterstützt Obsidian/Wiki- und Standard-Markdown-Bildeinbettung mit optionaler Größenangabe (px):

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- Ohne Größe: Originalgröße, begrenzt durch verfügbare Seiten-/Zellenbreite
- Mit Größe: Vergrößerung erlaubt, weiterhin durch verfügbare Breite begrenzt
- Nur-Bild-Absatz: verwendet den Absatzstil „Image“ (zentriert, einfacher Zeilenabstand)

## 📖 Detaillierte Bedienungsanleitung

### Word zu Markdown

1.  Ziehen Sie die `.docx`-Datei in das Programmfenster.
2.  Das Programm analysiert automatisch die Dokumentstruktur.
3.  Generiert eine `.md`-Datei mit YAML-Metadaten.

**Unterstützte Formate**:
-   `.docx` - Standard Word-Dokument.
-   `.doc` - Automatisch in DOCX zur Verarbeitung konvertiert.
-   `.wps` - WPS-Dokument automatisch konvertiert.

**Exportoptionen**:

| Option | Beschreibung |
| :--- | :--- |
| **Bilder extrahieren** | Wenn aktiviert, werden Bilder im Dokument in den Ausgabeordner extrahiert und Bildlinks in die MD-Datei eingefügt. |
| **Bild-OCR** | Wenn aktiviert, wird OCR auf Bildern durchgeführt und eine Bild-.md-Datei erstellt (mit erkanntem Text). |
| **Untertitelnummern bereinigen** | Wenn aktiviert, werden Nummern vor Untertiteln entfernt (z.B. "一、", "（一）", "1." usw.) und in reinen Titeltext konvertiert. |
| **Untertitelnummern hinzufügen** | Wenn aktiviert, werden Nummern basierend auf Überschriftenebenen automatisch hinzugefügt (Nummerierungsschema kann in den Einstellungen konfiguriert werden). |

### Markdown zu Word

1.  Bereiten Sie eine `.md`-Datei mit einem YAML-Header vor.
2.  Ziehen Sie sie in das Programmfenster und wählen Sie die entsprechende Word-Vorlage.
3.  Das Programm füllt die Vorlage automatisch und generiert das Dokument.

**Konvertierungsoptionen**:

| Option | Beschreibung |
| :--- | :--- |
| **Untertitelnummern bereinigen** | Wenn aktiviert, werden Nummern vor Untertiteln entfernt. |
| **Untertitelnummern hinzufügen** | Wenn aktiviert, werden Nummern basierend auf Überschriftenebenen automatisch hinzugefügt. |

**Hinweis**: Wenn es Absätze gibt, in denen Untertitel und Fließtext im Dokument gemischt sind, müssen strenge Zeilenumbrüche in der MD-Datei eingehalten werden (siehe "Zeilenumbrüche und Absätze" oben).

### Automatische Vorlagenstilverarbeitung

Der Konverter erkennt und verarbeitet Vorlagenstile automatisch während der Markdown → DOCX Konvertierung:

#### Stilklassifizierung

**Absatzstil**: Wird auf den gesamten Absatz angewendet.

| Stil | Erkennungsverhalten | Injektion bei Fehlen | Quelle |
| :--- | :--- | :--- | :--- |
| Überschrift (1~9) | Erkennt Absatzstil | Vorlagen-Überschriftsstile | Word Integriert |
| Codeblock | Erkennt Absatzstil | Consolas Schriftart + Grauer Hintergrund | Definiert durch Software |
| Zitat (1~9) | Erkennt Absatzstil | Grauer Hintergrund + Linker Rahmen | Definiert durch Software |
| Formelblock | Erkennt Absatzstil | Formelspezifischer Stil | Definiert durch Software |
| Trennzeichen (1~3) | Erkennt Absatzstil | Unterer Rahmen Absatzstil | Definiert durch Software |

**Zeichenstil**: Wird auf ausgewählten Text angewendet.

| Stil | Erkennungsverhalten | Injektion bei Fehlen | Quelle |
| :--- | :--- | :--- | :--- |
| Inline-Code | Erkennt Zeichenstil | Consolas Schriftart + Graue Schattierung | Definiert durch Software |
| Inline-Formel | Erkennt Zeichenstil | Formelspezifischer Stil | Definiert durch Software |

**Tabellenstil**: Wird auf die gesamte Tabelle angewendet.

| Stil | Erkennungsverhalten | Injektion bei Fehlen | Quelle |
| :--- | :--- | :--- | :--- |
| Dreilinientabelle | Benutzerkonfigurationspriorität | Dreilinientabellenstil-Definition | Definiert durch Software |
| Gittertabelle | Benutzerkonfigurationspriorität | Gittertabellenstil-Definition | Definiert durch Software |

**Nummerierungsdefinition**: Wird für Listenformate verwendet.

| Typ | Erkennungsverhalten | Behandlung bei Fehlen |
| :--- | :--- | :--- |
| Listennummerierung | Scannt vorhandene geordnete/ungeordnete Listendefinitionen in der Vorlage | Verwendet dezimal/bullet Voreinstellung |

#### Stilnamen-Internationalisierung

-   **Word Integrierte Stile** (heading 1~9):
    -   Stilnamen verwenden Word-Standard-englische Namen (z.B. `heading 1`).
    -   Word zeigt automatisch lokalisierte Namen basierend auf der Systemsprache an (z.B. "Überschrift 1" auf deutschen Systemen).
-   **Softwaredefinierte Stile** (Codeblock, Zitat, Formel, Trennzeichen, Tabelle usw.):
    -   Injiziert entsprechende Sprachstilnamen basierend auf der Schnittstellenspracheneinstellung der Software.
    -   Chinesische Schnittstelle: Injiziert "代码块", "引用 1", "三线表", usw.
    -   Englische Schnittstelle: Injiziert "Code Block", "Quote 1", "Three Line Table", usw.

**Vorschlag**: Nach dem Anpassen von Stilen in der Vorlage verwendet der Konverter automatisch Ihre Stile; wenn sie nicht in der Vorlage vorhanden sind, werden integrierte voreingestellte Stile verwendet.

### Tabellendateiverarbeitung

1.  **Excel/CSV zu Markdown**: Ziehen Sie `.xlsx` oder `.csv` Dateien, um sie automatisch in Markdown-Tabellen zu konvertieren.
2.  **Markdown zu Excel**: Bereiten Sie eine MD-Datei vor und wählen Sie eine Excel-Vorlage für die Konvertierung.

**Unterstützte Formate**:
-   `.xlsx` - Standard Excel-Dokument.
-   `.xls` - Automatisch in XLSX zur Verarbeitung konvertiert.
-   `.et` - WPS-Tabelle automatisch konvertiert.
-   `.csv` - CSV-Texttabelle.

### Textkorrekturfunktion

Das Programm bietet vier anpassbare Korrekturregeln:

1.  **Zeichenpaarungsprüfung** - Erkennt, ob gepaarte Satzzeichen wie Klammern und Anführungszeichen übereinstimmen.
2.  **Symbolkorrektur** - Erkennt gemischte Verwendung von chinesischen und englischen Satzzeichen.
3.  **Tippfehlerprüfung** - Überprüft auf häufige Tippfehler basierend auf einem benutzerdefinierten Wörterbuch.
4.  **Sensitive Worterkennung** - Erkennt sensitive Wörter basierend auf einem benutzerdefinierten Wörterbuch.

**Benutzerdefinierte Wörterbücher**: Bearbeiten Sie Tippfehler- und sensitive Wörterbücher visuell in der "Einstellungen"-Oberfläche.

**Verwendung**:
1.  Ziehen Sie das zu prüfende Word-Dokument in das Programm.
2.  Wählen Sie die erforderlichen Korrekturregeln.
3.  Klicken Sie auf die Schaltfläche "Textkorrektur".
4.  Korrekturergebnisse werden als Kommentare im Dokument angezeigt.

## 🛠️ Vorlagensystem

### Verwendung vorhandener Vorlagen

Das Programm enthält verschiedene Vorlagen, einschließlich mehrsprachiger Versionen. Sie können sie nach Bedarf auswählen und verwenden. Vorlagendateien befinden sich im Verzeichnis `templates/`.

### Benutzerdefinierte Vorlagen

1.  Erstellen Sie eine Vorlagendatei mit Word oder WPS.
2.  Beziehen Sie sich auf vorhandene Vorlagen und fügen Sie Platzhalter wie `{{Title}}` usw. ein, wo das Ausfüllen erforderlich ist.
3.  In der Vorlage müssen die integrierten Stile Überschrift 1 ~ Überschrift 5 manuell geändert werden.
4.  Speichern Sie die Vorlage im Verzeichnis `templates/`.
5.  Starten Sie das Programm neu, und die neue Vorlage wird automatisch geladen.

Sie können auch eine vorhandene Vorlage kopieren, ändern und umbenennen.

### Platzhalterverwendung

#### Word-Vorlagenplatzhalter

**YAML-Feldplatzhalter**: Verwenden Sie das Format `{{Feldname}}` in der Vorlage, das während der Konvertierung durch den entsprechenden Wert im YAML-Header der Markdown-Datei ersetzt wird.

| Platzhalter | Beschreibung |
| :--- | :--- |
| `{{Titel}}` | Dokumenttitel (Abrufregeln siehe unten) |
| `{{Inhalt}}` | Einfügeposition für Markdown-Textkörper |
| Andere | Unterstützt jedes benutzerdefinierte Feld |

**Titelabruf-Priorität**:

| Priorität | Quelle | Beschreibung |
| :--- | :--- | :--- |
| 1 | YAML `Title` Feld | Höchste Priorität |
| 2 | YAML `aliases` Feld | Nimmt das erste Element der Liste oder den Zeichenfolgenwert |
| 3 | Dateiname | Dateiname ohne `.md` Erweiterung |

**Mehrsprachige Unterstützung**: Die Platzhalter für Titel und Inhalt unterstützen mehrere Sprachen, z.B. Titel kann `{{Titel}}`, `{{title}}`, `{{标题}}` usw. sein, Inhalt kann `{{Inhalt}}`, `{{body}}`, `{{正文}}` usw. sein.

#### Excel-Vorlagenplatzhalter

Excel-Vorlagen unterstützen drei Arten von Platzhaltern:

**1. YAML-Feldplatzhalter** `{{Feldname}}`

Wird verwendet, um einen einzelnen Wert aus dem YAML-Header der Markdown-Datei auszufüllen:

```markdown
---
ReportName: 2024 Jahresverkaufsstatistik
Unit: Verkaufsabteilung
---
```

`{{ReportName}}`, `{{Unit}}` in der Vorlage werden durch entsprechende Werte ersetzt. Das Titelfeld folgt ebenfalls den Prioritätsregeln.

**2. Spaltenfüllplatzhalter** `{{↓Feldname}}`

Extrahiert Daten aus der Markdown-Tabelle und füllt **nach unten** Zeile für Zeile ab der Platzhalterposition:

```markdown
| Produktname | Menge |
|:--- |:--- |
| Produkt A | 100 |
| Produkt B | 200 |
```

`{{↓Produktname}}` in der Excel-Vorlage wird durch "Produkt A" ersetzt, und die nächste Zeile wird mit "Produkt B" gefüllt.

**3. Zeilenfüllplatzhalter** `{{→Feldname}}`

Extrahiert Daten aus der Markdown-Tabelle und füllt **nach rechts** Spalte für Spalte ab der Platzhalterposition:

```markdown
| Monat |
|:--- |
| Jan |
| Feb |
| Mär |
```

`{{→Monat}}` in der Excel-Vorlage wird nacheinander mit "Jan", "Feb", "Mär" nach rechts gefüllt.

**Behandlung verbundener Zellen**: Das Programm überspringt automatisch nicht erste Zellen verbundener Zellen, um eine korrekte Dateneingabe zu gewährleisten.

**Zusammenführung mehrerer Tabellendaten**: Wenn es mehrere Tabellen in Markdown gibt, die denselben Kopfnamen verwenden, werden die Daten der Reihe nach zusammengeführt und nacheinander gefüllt.

## 🖥️ Verwendung der grafischen Oberfläche

Die meisten Benutzer verwenden diese Software über die grafische Oberfläche. Hier ist die detaillierte Bedienungsanleitung.

### Schnittstellenübersicht

Das Programm verwendet ein **adaptives dreispaltiges Layout**:

| Bereich | Beschreibung | Anzeigezeitpunkt |
| :--- | :--- | :--- |
| **Mittlere Spalte (Hauptbereich)** | Datei-Drag-and-Drop-Bereich, Bedienfeld, Statusleiste | Immer angezeigt |
| **Rechte Spalte** | Vorlagenauswahl / Formatkonvertierungspanel | Erweitert sich automatisch nach Auswahl einer Datei |
| **Linke Spalte** | Stapeldateiliste (gruppiert nach Typ) | Wird angezeigt, wenn in den Stapelmodus gewechselt wird |

### Grundlegender Arbeitsablauf

1.  **Programm starten**: Doppelklicken Sie auf `DocWen.exe`.
2.  **Datei importieren**:
    -   Methode 1: Ziehen Sie Dateien direkt in das Fenster.
    -   Methode 2: Klicken Sie auf die Schaltfläche "Hinzufügen" im Drag-and-Drop-Bereich, um Dateien auszuwählen.
3.  **Vorlage auswählen** (falls Konvertierung erforderlich): Das rechte Vorlagenpanel erweitert sich automatisch; wählen Sie eine geeignete Vorlage.
4.  **Optionen konfigurieren**: Wählen Sie im Bedienfeld die erforderlichen Konvertierungs-/Exportoptionen aus.
5.  **Operation ausführen**: Klicken Sie auf die entsprechende Funktionstaste (z.B. "Export MD", "In DOCX konvertieren" usw.).
6.  **Ergebnis anzeigen**: Die Statusleiste zeigt Fortschritt und Ergebnisse an; klicken Sie auf das 📍-Symbol, um die Ausgabedatei zu finden.

### Einzeldateimodus vs. Stapelmodus

Das Programm unterstützt zwei Verarbeitungsmodi, die über die Umschalttaste im Datei-Drag-and-Drop-Bereich umgeschaltet werden können:

**Einzeldateimodus** (Standard):
-   Verarbeitet jeweils eine Datei.
-   Einfache Schnittstelle, geeignet für den täglichen Gebrauch.

**Stapelmodus**:
-   Importiert mehrere Dateien gleichzeitig.
-   Linke Spalte zeigt kategorisierte Dateiliste (gruppiert nach Dokument/Tabelle/Bild usw.).
-   Unterstützt Stapelhinzufügen, Entfernen und Sortieren.
-   Klicken auf eine Datei in der Liste wechselt das aktuelle Operationsziel.

### Bedienfeldfunktionen

Das Bedienfeld passt die verfügbaren Optionen automatisch basierend auf dem Dateityp an:

| Dateityp | Verfügbare Operationen |
| :--- | :--- |
| Word-Dokument | Export MD, Konvertieren PDF, Textkorrektur, OCR |
| Markdown | Konvertieren DOCX, Konvertieren PDF |
| Excel-Tabelle | Export MD, Konvertieren PDF, Tabellenzusammenfassung |
| PDF-Datei | Export MD, Zusammenführen, Teilen, OCR |
| Bilddatei | Formatkonvertierung, Komprimierung, OCR |

### Einstellungsoberfläche

Klicken Sie auf die ⚙️-Schaltfläche in der unteren rechten Ecke des Fensters, um die Einstellungen zu öffnen:

-   **Allgemein**: Schnittstellendesign, Sprache, Fensterdeckkraft.
-   **Konvertierung**: Standardwerte für verschiedene Konvertierungsoptionen.
-   **Ausgabe**: Standardausgabeverzeichnis, Dateibenennungsregeln.
-   **Korrektur**: Bearbeiten Sie Tippfehler- und sensitive Wörterbücher.
-   **Stil**: Codeblock-, Zitat-, Tabellenstilkonfigurationen.

### Verknüpfungen

-   **Externe Datei ziehen**: Ziehen Sie direkt in das Fenster zum Importieren.
-   **Doppelklick auf Statusleistenergebnis**: Öffnen Sie schnell das Ausgabedateiverzeichnis.
-   **Rechtsklick auf Vorlagenelement**: Öffnen Sie den Vorlagendateispeicherort.

---

## 🔧 Befehlszeilennutzung

Zusätzlich zur GUI bietet das Programm eine Befehlszeilenschnittstelle (CLI), die für Automatisierungsskripte und Stapelverarbeitungsszenarien geeignet ist.

### Ausführungsmodi

-   **Interaktiver Modus**: Zeigt eine Menüführung nach dem Übergeben einer Datei an, ähnlich wie bei der GUI-Bedienung.
-   **Headless-Modus**: Direkt durch Hinzufügen des `--action` Parameters ausführen, geeignet für Skriptaufrufe.

### Häufige Beispiele

```bash
# Interaktiver Modus
DocWen.exe document.docx

# Word in Markdown exportieren (Bilder extrahieren + OCR)
DocWen.exe report.docx --action export_md --extract-img --ocr

# Markdown zu Word (Vorlage angeben)
DocWen.exe document.md --action convert --target docx --template "Vorlagenname"

# Stapelkonvertierung (Bestätigung überspringen, bei Fehler fortfahren)
DocWen.exe *.docx --action export_md --batch --yes --continue-on-error

# Dokumentenprüfung
DocWen.exe document.docx --action validate --check-typo --check-punct

# PDF Zusammenführen/Teilen
DocWen.exe *.pdf --action merge_pdfs
DocWen.exe report.pdf --action split_pdf --pages "1-3,5,7-10"
```

### Hauptargumente

| Argument | Beschreibung |
| :--- | :--- |
| `--action` | Operationstyp: `export_md`, `convert`, `validate`, `merge_pdfs`, `split_pdf` |
| `--target` | Zielformat: `pdf`, `docx`, `xlsx`, `md` |
| `--template` | Vorlagenname (z.B. `Vorlagenname`) |
| `--extract-img` | Bilder beim Export extrahieren |
| `--ocr` | OCR-Erkennung aktivieren |
| `--batch` | Stapelverarbeitungsmodus |
| `--yes` / `-y` | Bestätigungsaufforderungen überspringen |
| `--continue-on-error` | Bei Fehler mit dem nächsten Element fortfahren |
| `--json` | Ergebnis im JSON-Format ausgeben |
| `--quiet` / `-q` | Stiller Modus, Ausgabe reduzieren |

## 🔌 Obsidian-Plugin

Das Projekt enthält ein passendes Obsidian-Plugin, um mit dem Konverter zusammenzuarbeiten:

### Kernfunktionen

-   **🚀 Ein-Klick-Start** - Seitenleistensymbol zum schnellen Starten des Konverters.
-   **� Automatische Übergabe** - Übergibt automatisch den aktuell geöffneten Dateipfad.
-   **🔄 Einzelinstanzverwaltung** - Sendet Datei automatisch, wenn das Programm bereits läuft, kein Neustart erforderlich.
-   **💪 Absturzwiederherstellung** - Erkennt automatisch den Prozessstatus und bereinigt automatisch verbleibende Dateien.

### Funktionsprinzip

Das Plugin interagiert mit dem Konverter über dateisystembasierte IPC:

1.  **Erster Klick** → Konverter starten und aktuelle Datei übergeben.
2.  **Klick erneut (Mit Datei)** → Durch neue Datei ersetzen (Einzeldateimodus).
3.  **Klick erneut (Keine Datei)** → Konverterfenster aktivieren.

### Installation

Das Plugin wurde in einem separaten Repository veröffentlicht. Besuchen Sie bitte [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) für Installationsanweisungen und die neueste Version.

## ❓ FAQ

### Was tun, wenn die Konvertierung fehlschlägt?

-   Prüfen Sie, ob die Datei von einem anderen Programm belegt ist.
-   Bestätigen Sie, dass das Dateiformat korrekt ist.
-   Überprüfen Sie Fehlerprotokolle im Verzeichnis `logs/`.

### Vorlage wird nicht angezeigt?

-   Bestätigen Sie, dass sich Vorlagendateien im Verzeichnis `templates/` befinden.
-   Prüfen Sie, ob die Vorlagendatei beschädigt ist.
-   Starten Sie das Programm neu, um Vorlagen neu zu laden.

### Korrekturfunktion funktioniert nicht?

-   Bestätigen Sie, dass das Dokument im .docx-Format vorliegt.
-   Prüfen Sie, ob das Dokument bearbeitbaren Text enthält.
-   Bestätigen Sie, dass Korrekturregeln in den Einstellungen aktiviert sind.

### Ausgabeformat nicht wie erwartet?

-   Das Programm generiert Dokumente basierend auf Vorlagenstilen. Um das Ausgabeformat anzupassen, ändern Sie die Stildefinitionen direkt in der Vorlagendatei.
-   Vorlagendateien befinden sich im Verzeichnis `templates/`.
-   Nach Änderung der Vorlagenstile werden alle mit dieser Vorlage konvertierten Dokumente die neuen Stile anwenden.

### Formelzellen sind nach der Excel-zu-Markdown-Konvertierung leer?

Dies ist ein erwartetes Verhalten. Das Programm liest die **zwischengespeicherten Werte** der Zellen anstelle der Formeln selbst.

**Technischer Grund**:
-   In Excel-Dateien speichern Formelzellen sowohl die Formel als auch das zuletzt berechnete Ergebnis (zwischengespeicherter Wert).
-   Das Programm verwendet den Modus `data_only=True`, der nur zwischengespeicherte Werte abruft.
-   Wenn die Datei nie in Excel geöffnet wurde (z.B. von einem Programm generiert) oder bearbeitet, aber nicht erneut gespeichert wurde, ist der zwischengespeicherte Wert leer.

**Lösung**:
1.  Öffnen Sie die Datei in Excel.
2.  Warten Sie, bis die Formelberechnung abgeschlossen ist.
3.  Speichern Sie die Datei.
4.  Konvertieren Sie erneut.

## 🔒 Sicherheitsfunktionen

-   **Vollständig lokaler Betrieb**: Alle Verarbeitungen erfolgen lokal, keine Netzwerkabhängigkeit.
-   **Netzwerkisolation**: Eingebauter Netzwerkisolationsmechanismus verhindert Datenlecks.
-   **Kein Daten-Upload**: Benutzerdateien werden niemals auf einen Server hochgeladen.

## 📜 Lizenz

Dieses Projekt ist unter der **GNU Affero General Public License v3.0 (AGPL-3.0)** lizenziert.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Dieses Projekt verwendet PyMuPDF (lizenziert unter AGPL-3.0), daher ist das gesamte Projekt ebenfalls unter AGPL-3.0 lizenziert.
-   Sie dürfen diese Software frei verwenden, ändern und verbreiten.
-   Wenn Sie diese Software ändern und Dienste über ein Netzwerk anbieten, müssen Sie den Benutzern den geänderten Quellcode zur Verfügung stellen.
-   Detaillierte Lizenzinformationen finden Sie in der Datei [LICENSE](LICENSE).

### Kontakt

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Autor kontaktieren**: zhengyx91@hotmail.com

---

**Autor**: ZhengYX
