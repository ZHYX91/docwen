# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md) · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)

DocWen ist ein lokales Werkzeug zur Konvertierung von Dokumenten und Tabellen für Word, Markdown, Excel und verwandte Büroformate. Es ist für zuverlässige Offline-Arbeitsabläufe ausgelegt und verarbeitet Dokumente auf dem Gerät des Benutzers.

Die Oberflächensprache steuert Texte und Anordnung der Anwendung. Die separate Ausgabesprache in den allgemeinen Einstellungen steuert erzeugte Beschriftungen, Dokumentstile und Korrekturkommentare, ohne den Inhalt zu übersetzen. Ein Sprachwechsel ändert weder die gespeicherte Ausgabesprache noch die Vorlage. OCR „Automatisch“ folgt der Ausgabesprache. In den Markdown-Einstellungen kann jede gewöhnliche Textzeile optional einen eigenen Word-Absatz bilden; die Option ist standardmäßig aus.

## 📖 Projekthintergrund

DocWen entstand für alltägliche Dokumentaufbereitung, bei der typischerweise:
- Dokumente aus verschiedenen Abteilungen in uneinheitlichen Formaten vereinheitlicht werden müssen;
- viele Dateitypen mit unterschiedlichen Formatvorgaben vorkommen;
- die Arbeit offline, im Intranet und auf älterer Hardware möglich sein soll.

**Designphilosophie**: DocWen setzt auf leicht zugängliche, sofort nutzbare Abläufe für alltägliche Konvertierungs- und Aufräumarbeiten. Es soll spezialisierte Satz- oder Konvertierungssysteme wie LaTeX oder Pandoc nicht ersetzen, sondern häufige Dokumentaufgaben verständlich und direkt machen und dabei klare technische Grenzen beibehalten.

## ✨ Kernfunktionen

- **📄 Dokumentformatkonvertierung** - Bidirektionale Word ↔ Markdown Konvertierung. Unterstützt mathematische Formelkonvertierung, bidirektionale Trennzeichenkonvertierung (Markdowns drei Arten von Trennlinien vs. Words Seitenumbrüche, Abschnittswechsel und horizontale Linien) sowie die Wiederherstellung expliziter Markdown-Tabellenmarker `<` / `^` zu rechteckigen Word-Zellzusammenführungen. Unterstützt Formate wie DOCX/DOC/WPS/RTF/ODT.
- **📊 Tabellenkonvertierung** – Excel ↔ Markdown für XLSX/XLS/ET/ODS/CSV/TSV. Zusammengeführte Zellen lassen sich beim Export ausfüllen, leer lassen oder mit Markierungen erhalten. XLSX-Vorlagen unterstützen YAML-Felder, Tabellenzeilen und -spalten, Bilder und zusammengeführte Zellen; mehrere Tabellen können zusammengefasst werden.
- **📑 PDF und Layoutdateien** - PDF/XPS/OFD zu Markdown oder DOCX Konvertierung. Unterstützt PDF-Zusammenführung, -Teilung und andere Operationen.
- **🖼️ Bildverarbeitung** - Unterstützt bidirektionale Konvertierung und Komprimierung von JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC Formaten.
- **📥 Import anderer Formate** - Unterstützt die einseitige Konvertierung von HTML/MHTML/ENEX/EPUB/PPTX/PPT nach Markdown.
- **🔍 OCR-Texterkennung** - Integriertes RapidOCR zum Extrahieren von Text aus Bildern und PDFs.
- **✏️ Textkorrektur** - Überprüft Word (.docx) und Markdown (.md) Dateien auf Tippfehler, Zeichensetzung, Symbole und sensitive Wörter basierend auf benutzerdefinierten Wörterbüchern. Regeln können in der Einstellungsoberfläche bearbeitet werden.
- **📝 Vorlagensystem** - Flexibler Vorlagenmechanismus, der benutzerdefinierte Dokument- und Berichtsformate unterstützt.
- **💻 Dual-Modus-Betrieb** - Grafische Benutzeroberfläche (GUI) + Befehlszeilenschnittstelle (CLI).
- **🔒 Lokale Verarbeitung mit Abfluss-Schutz für Abhängigkeiten** - Die Konvertierung benötigt keine Onlinedienste. Während DocWen läuft, blockiert der Python-Prozess DNS sowie IPv4/IPv6 für prozessinterne Abhängigkeiten; extern gestartete Office-Anwendungen folgen ihren eigenen System-Netzwerkregeln.
- **🔗 Einzelinstanzbetrieb** - Verwaltet automatisch Programminstanzen und unterstützt die Integration mit dem begleitenden Obsidian-Plugin.

## 📸 Screenshots

Vorlage und Erstellungsoptionen für Markdown-Dokumente auswählen.

![Markdown in Dokumente umwandeln](../assets/screenshots/markdown-generation-en.png)


Changelog: siehe [CHANGELOG.md](../CHANGELOG.md)

## 🚀 Schnellstart

### Installation aus Quellcode

**Voraussetzungen**: Python 3.12

**Paketziele**: Windows x64, Ubuntu 24.04 x64 und Apple Silicon macOS 14+ arm64. Keine Intel-Mac-Pakete.

**Option 1: Mit uv (empfohlen)**

Installieren Sie [uv](https://docs.astral.sh/uv/getting-started/), dann:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

Quellcode, Tests und Builds von DocWen unterstützen nur die eingecheckte Sperrdatei mit `uv 0.12.0`; `pip install -e` wird nicht unterstützt.

### Programm starten

Für das Windows-Paket: Doppelklicken Sie auf `DocWen.exe`, um die GUI zu starten. Nach Installation aus Quellcode:

```bash
docwen-gui  # GUI-Modus
docwen      # CLI-Modus
```

### Hinweise für macOS

Mac-Pakete sind ad-hoc signiert, ohne Developer ID und Apple-Notarisierung. Siehe [Installation und erster Start](macos.md).

**LibreOffice-Unterstützung (Optional)**

Um ältere Formate wie `.doc` und `.xls` zu konvertieren, installieren Sie LibreOffice:  
Download: https://www.libreoffice.org/download/

**HEIC-Bildunterstützung (Optional)**

Um HEIC/HEIF-Bilder zu verarbeiten:

```bash
brew install libheif
pip install pillow-heif
```

### Voraussetzungen für die Linux-GUI

**Unterstütztes Paketziel**: DocWen unterstützt GUI und CLI im Ubuntu-24.04-x64-Paket. Diese
Voraussetzungen erweitern die Zusage nicht auf andere Distributionen oder Architekturen.

- Eine Desktop-Umgebung ist installiert (GNOME, KDE, XFCE usw.)
- Die GUI basiert auf PySide6 (Qt6) und hängt nicht mehr von Python Tk ab. Falls beim Start Systembibliotheken fehlen, installieren Sie die im Fehler genannten Qt-Laufzeitabhängigkeiten (häufig OpenGL/X11-bezogen).
- Auf headless Servern sollten Sie den CLI-Einstieg `docwen` statt der GUI verwenden; in Windows-Paketen steht zusätzlich `DocWenCLI.exe` bereit.

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
    - Klicken Sie auf "Generieren DOCX".

3.  **Ergebnis erhalten**:
    - Das Word-Dokument wird in einem Ergebnisordner am eingestellten Ausgabeort gespeichert. Öffnen Sie den Ausgabeort über die Statusleiste.

**Tipp**: Sie können die Beispieldateien im Verzeichnis `samples/` verwenden, um die Funktionen der Software schnell auszuprobieren.

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

1.  **Programm starten**: Doppelklicken Sie auf `DocWen.exe` (Windows-Paket) oder führen Sie `docwen-gui` aus.
2.  **Datei importieren**:
    -   Methode 1: Ziehen Sie Dateien direkt in das Fenster.
    -   Methode 2: Klicken Sie auf die Schaltfläche "Hinzufügen" im Drag-and-Drop-Bereich, um Dateien auszuwählen.
3.  **Vorlage auswählen** (falls Konvertierung erforderlich): Das rechte Vorlagenpanel erweitert sich automatisch; wählen Sie eine geeignete Vorlage.
4.  **Optionen konfigurieren**: Wählen Sie im Bedienfeld die erforderlichen Konvertierungs-/Exportoptionen aus.
5.  **Operation ausführen**: Klicken Sie auf die entsprechende Funktionstaste (z.B. "Markdown Export", "Generieren DOCX" usw.).
6.  **Ergebnis anzeigen**: Die Statusleiste zeigt Fortschritt und Ergebnisse an; klicken Sie rechts auf die Aktion „Ausgabe öffnen“, um den Ausgabeort zu öffnen.

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
| Word-Dokument | Markdown Export, Konvertieren PDF, Textkorrektur, OCR |
| Markdown | Konvertieren DOCX, Konvertieren PDF, Textkorrektur |
| Excel-Tabelle | Markdown Export, Konvertieren PDF, Tabellenzusammenfassung |
| PDF-Datei | Markdown Export, Zusammenführen, Teilen, OCR |
| Bilddatei | Formatkonvertierung, Komprimierung, OCR |
| HTML/EPUB/PPTX usw. | Markdown Export |

### Einstellungsoberfläche

Öffnen Sie die Einstellungen über das Symbol am unteren Fensterrand oder mit **Strg+,**. Die Navigation ist in **Allgemein**, **Eingabetypen**, **Inhalte** und **Spezialfunktionen** gegliedert. In niedrigen Fenstern ist die Seitenleiste scrollbar; bei zu geringer Breite erscheint eine Seitenauswahl.

**OCR und Bilder** enthält gemeinsame Optionen für Erkennung und Markdown-Bildausgabe. Die OCR-Vorgabe eines Dateityps bleibt auf dessen Eingabeseite; Bildkonvertierung und Komprimierung stehen unter **Bildeingabe**. **Markdown-Syntax** trennt allgemeine Ausgaberegeln, Eingaberegeln und Word-Zuordnungen. Nicht jeder Eingabeweg oder jede Spezialoptimierung berücksichtigt schon alle Syntaxoptionen.

**Amtliche Dokumente** und **Rechnungen** haben eigene Seiten. Bei amtlichen Dokumenten regelt **Dokument → Markdown** die Erkennungsvorgabe und **Markdown → Dokument** die Formatierung von Vorlagenfeldern mit Hinweisen zu Anlagen, Verteiler, Datum und Vermerken. Rechnungen bieten getrennte Vorgaben für **PDF / OFD → Markdown** und **Bild → Markdown**. Diese Optionen bestimmen nur die anfängliche Auswahl geeigneter Dateien und klassifizieren sie nicht automatisch. **Übernehmen / OK** speichert, **Abbrechen** verwirft ungespeicherte Änderungen.

**Geltungsbereich:** Die spezielle Feldformatierung ist für chinesische Amtsdokumente ausgelegt; die Rechnungsoptimierung zielt auf chinesische Rechnungsformate. Andere Dateien können weiterhin die normalen Konvertierungswege verwenden.

### Verknüpfungen

-   **Externe Datei ziehen**: Ziehen Sie direkt in das Fenster zum Importieren.
-   **Ausgabe öffnen**: Klicken Sie rechts in der Statusleiste auf die Aktion „Ausgabe öffnen“, um den Ausgabeort zu öffnen.
-   **Rechtsklick auf Vorlagenelement**: Öffnen Sie den Vorlagendateispeicherort.

---

## 🔧 CLI-Verwendung

Die Kommandozeile (CLI) eignet sich zum Prüfen und Konvertieren von Dateien im Terminal sowie für lokale Automatisierungsskripte. Für die Einbindung in andere Apps verwenden Sie die unten beschriebene Machine-Protocol-Schnittstelle.

### Empfohlener Ablauf für Automatisierung

Für lokale Automatisierungsskripte können Sie die Konvertierung in dieser Reihenfolge prüfen und ausführen:

1. `inspect <file> [--json]`: zuerst den tatsächlichen Dateityp, das Format und die unterstützten Aktionen erkennen.
2. `resources list formats --json`: die in DocWen tatsächlich verfügbaren Routen und Abhängigkeitsbedingungen lesen.
3. `schema convert`: den maschinenlesbaren Vertrag und die Bedingungen von `convert` abrufen.
4. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) --dry-run --json`: Erkennung, Normalisierung und Routing vorab prüfen, ohne Dateien zu schreiben.
5. `convert <file> --to <fmt> (--output-dir <dir> | --output <path>) ...`: die echte Konvertierung erst danach ausführen.

Apps und Plugins verbinden sich über `serve --stdio` mit Machine Protocol v2, um verfügbare Aktionen abzufragen, Aufgaben zu planen und sie auszuführen. Dies ist die stabile Integrationsschnittstelle; CLI `--json` ist für Terminalskripte gedacht. Mit `gui open|activate|status` öffnen Sie Dateien in der Desktop-App oder holen ihr Fenster in den Vordergrund.

### Häufige Beispiele

Konvertierungen von oder nach Markdown erstellen auch bei nur einer Ausgabedatei einen Ergebnisordner unter `--output-dir DIR`. Namen enthalten Quellname, Zeitstempel und aktuelles Eingabeformat. `--output PATH` ist anderen Konvertierungen mit genau einer Zieldatei vorbehalten.

```bash
# Windows-Paket
DocWenCLI.exe inspect document.docx --json

# Run-Vertrag für Skripte / Agents exportieren
DocWenCLI.exe schema convert

# Ablauf der Konvertierung prüfen, ohne Dateien zu erzeugen
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr --dry-run --json

# Word nach Markdown exportieren (Bilder extrahieren + OCR)
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --ocr

# Markdown nach Word (Vorlage und Überschrift/Text-Zusammenführung)
DocWenCLI.exe convert document.md --to docx --output-dir exports --template template.docx.e26c3c7ebfb4e8bb1118e913afe80ace9ff4371da48bc1ba988e4b7578c609df --heading-merge-mode punct_required

# Markdown nach Word und erzeugte DOCX-Datei vor der Veröffentlichung prüfen
DocWenCLI.exe convert document.md --to docx --output-dir exports --proofread --check typo --check punct

# Bildmodus und OCR-Platzierung für Markdown steuern
DocWenCLI.exe convert report.docx --to md --output-dir exports --extract-img --image-mode file --ocr --ocr-placement image_md

# Laufzeitfähigkeiten und Abhängigkeits-Gates anzeigen
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Dokumentprüfung
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# Aus dem Quellcode / per uv
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output-dir exports --dry-run --json
# docwen convert document.docx --to md --output-dir exports
```

### Haeufige Befehle und Optionen

Die folgende Tabelle zeigt nur haeufige Befehle. Fuer die vollstaendige Befehlsflaeche verwenden Sie `docwen --help` (Quellcode / uv) oder `DocWenCLI --help` (Paketversion).

| Befehl / Option | Beschreibung |
| --- | --- |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>)` | Einheitlicher Einstiegspunkt für Konvertierungen. |
| `convert <markdown> --to docx --output-dir <dir> --proofread [--check ...]` | Korrektur nach Markdown→DOCX; `--check` ist bei `convert` nur zusammen mit `--proofread` gültig. Ohne `--check` gelten die konfigurierten Korrektur-Standardwerte. |
| `convert <file> --to <fmt> (--output-dir <dir> \| --output <path>) --dry-run --json` | Prüft Erkennung, Normalisierung, Routing und wirksame Optionen, ohne die Konvertierung wirklich auszuführen. |
| `schema convert` | Exportiert den maschinenlesbaren Vertrag, Standardwerte, Bedingungen und die kanonischen Schlüssel von `convert`. |
| `validate <file> --check ...` | Dokumentprüfung (`typo/punct/symbol/sensitive/all/none`). Verwenden Sie `--json` für die CLI-Hülle; `--report` ist ein optionaler Bericht-Dateipfad. |
| `inspect <file> [--json]` | Dateikategorie/-format, empfohlene Aktionen und Warnungen bei Erweiterungs-/Inhaltsabweichungen anzeigen. |
| `doctor --json` | Gibt Diagnosen zusammen mit Laufzeitfähigkeits-Zusammenfassungen und Abhängigkeits-Gates aus. |
| `resources list formats --json` | Listet Zielformate nach Quellkategorie auf und ergänzt Zusammenfassungen zu Laufzeit-Abhängigkeiten und Einschränkungen. |
| `resources list optimizations --json` | Typisierte Optimierungsressourcen und ihre kanonischen Routenbindungen auflisten. |
| `resources list templates` | Verfügbare Vorlagen auflisten. |
| `resources list numbering-schemes` | Verfügbare Nummerierungsschemata auflisten. |
| `gui open [ABSOLUTE_FILE] [--timeout <seconds>] [--json]` | DocWen über die getrennte lokale GUI-Steuerung starten oder aktivieren und optional eine absolute Datei öffnen. |
| `--template <id>` | Exakte kanonische Ressourcen-ID aus `resources list templates`; Anzeigenamen, Dateinamen und Pfade werden abgelehnt. DOCX-IDs gelten für `docx/doc/odt/rtf/wps/pdf`, XLSX-IDs für `xlsx/xls/ods/csv`. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Bildextraktion und OCR für `convert --to md`. |
| `--image-mode file\|base64\|embed\|omit` | Steuert, wie Bilder beim Markdown-Export ausgegeben werden. |
| `--ocr-placement image_md\|main_md` | Legt fest, ob OCR-Text in begleitendes Bild-Markdown oder in die Haupt-Markdown-Datei geschrieben wird. |
| `--heading-merge-mode punct_required\|always\|never` | Steuert die Strategie für die Zusammenführung von Überschrift + Text bei `convert --to docx`. |
| `--optimization <id>` | Aktiviert explizit ein Optimierungsprofil (siehe `resources list optimizations`). |
| `batch convert\|validate ... --jobs <n> [--continue-on-error]` | Steuerung der Stapelverarbeitung. |
| `--json` / `--quiet` / `--timing` | JSON ausgeben, Protokolle reduzieren oder die Laufzeit in Terminalskripten anzeigen. |

Im Modus `punct_required` lautet die genaue Standardliste `。：！？.:!?`. Sie kann in den Formatierungseinstellungen bearbeitet werden; ein leerer Wert deaktiviert die Zusammenführung in diesem Modus. Komma, Semikolon, Aufzählungskomma, Gedankenstrich und Auslassungspunkte sind standardmäßig ausgeschlossen.


## 📝 Markdown-Syntaxkonventionen

**Obsidian-Markdown-Erweiterungen verwenden?** Bei einer neuen DocWen-Installation sind vier optionale Erweiterungen **ausgeschaltet**: `structural_tables` (verbundene Zellen und mehrzeilige Tabellenköpfe), `captions_references` (Number-Suite-Beschriftungen und Querverweise), `extended_headings` (H7–H9) und `typed_endnotes` (typisierte Endnoten).

- **Markdown → Word:** Aktivieren Sie unter **Einstellungen → Markdown-Syntax → Eingabe** die benötigten Erweiterungen oder wählen Sie für die Eingabe **Obsidian-Erweiterungen verwenden**. Sonst können Marker als gewöhnlicher Text erhalten bleiben.
- **Word → Markdown:** Aktivieren Sie die entsprechenden **Ausgabe**-Erweiterungen gesondert oder nutzen Sie die Ausgabe-Voreinstellung; ansonsten können Strukturen vereinfacht oder normalisiert werden, mit Warnhinweisen.

Bewahren Sie das ursprüngliche Markdown auf und prüfen Sie die DOCX-Ausgabe. Word-Nummern folgen den DocWen-Konvertierungsoptionen; identische Schreibweisen und Leerzeichen werden bei der Rückkonvertierung nicht garantiert. Details: [Markdown-Kompatibilität](../specs/markdown-compatibility.md).

### Überschriftenebenen-Zuordnung

Um es Kollegen ohne Hintergrundwissen leichter zu machen, entsprechen die Markdown-Überschriften in dieser Software **eins-zu-eins** den Word-Überschriften:
- Dokumenttitel (title) und Untertitel (subtitle) werden in den YAML-Metadaten platziert.
- Markdown `# Überschrift 1` entspricht Word "Überschrift 1".
- Markdown `## Überschrift 2` entspricht Word "Überschrift 2".
- Und so weiter, bis zu 9 Überschriftenebenen werden unterstützt.

**Tipp**: Wenn Sie es bevorzugen, Markdowns erste Überschriftenebene (`#`) als Dokumenttitel zu verwenden und ab der zweiten Ebene (`##`) mit Textüberschriften zu beginnen, können Sie den Stil "Überschrift 1" in der Word-Vorlage so anpassen, dass er wie ein Dokumenttitel aussieht (z.B. zentriert, fett, größere Schriftgröße), und in den Einstellungen ein Nummerierungsschema auswählen, das die Nummerierung der ersten Überschriftenebene überspringt. So erscheinen Ihre Überschriften der ersten Ebene als Dokumenttitel.

### Zeilenumbrüche und Absätze

**Grundregel**: Normale weiche Markdown-Zeilenumbrüche erzeugen standardmäßig **keine** getrennten Word-Absätze. Aktivieren Sie **Jede Zeile als Absatz** unter **Einstellungen → Markdown-Syntax**, um gewöhnliche Textzeilen zu trennen. Listen, Zitate, Code und ausdrückliche harte Umbrüche behalten eigene Regeln.

**Gemischte Absätze**: Wenn ein Untertitel mit dem Fließtext im selben Absatz gemischt werden muss (Standardmodus: „Satzzeichen erforderlich“), müssen folgende Bedingungen erfüllt sein:
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
- Standardmäßig (Modus „Satzzeichen erforderlich“): Wenn der Untertitel nicht mit einem abschließenden Satzzeichen endet, wird er auch ohne Leerzeile nicht mit der nächsten Zeile zusammengeführt.
- Ändern Sie den Modus und die auslösenden Zeichen unter **Einstellungen → Texteingabe → Markdown zu Dokument — Formatverarbeitung → Modus zum Zusammenführen von Überschrift und Text**.

### Bidirektionale Trennzeichenkonvertierung

Unterstützt die bidirektionale Konvertierung zwischen Markdown-Trennzeichen und Word-Seitenumbrüchen/Abschnittswechseln/horizontalen Linien:

-   **DOCX → MD**: Word-Seitenumbrüche, Abschnittswechsel und horizontale Linien werden automatisch in Markdown-Trennzeichen konvertiert.
-   **MD → DOCX**: Markdown `---`, `***`, `___` werden automatisch in entsprechende Word-Elemente konvertiert.
-   **Konfigurierbar**: Spezifische Zuordnungsbeziehungen können in der Einstellungsoberfläche angepasst werden.

### Aufgabenlisten

Unterstützt die bidirektionale Konvertierung von GFM-Aufgabenlisten:

```markdown
- [ ] Aufgabe
- [x] Erledigt
```

-   **MD → DOCX**: Wird als Aufzählungsliste mit `☐` / `☑` Textpräfix gerendert.
-   **DOCX → MD**: Konvertiert Listenelemente mit `☐` / `☑` / `☒` Präfix zurück zu `- [ ]` / `- [x]`.
-   **Schriftart-Hinweis**: `☐`/`☑` werden möglicherweise in einigen Schriftarten nicht angezeigt. Verwenden Sie bei Bedarf Schriftarten wie „Segoe UI Symbol" in Ihrer Word-Vorlage.

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

### Link-Verarbeitung

Unterstützt klickbare Links bei Markdown -> DOCX:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Markdown-Links und Wiki-Links werden standardmäßig als Word-Hyperlinks ausgegeben
- Wiki-Links werden als lokale `file:///`-Links aufgelöst, wenn die Zieldatei gefunden wird
- Autolinks in spitzen Klammern unterstützen `https://...` und E-Mail-Links `mailto:...`
- Die automatische Verlinkung nackter URLs wird für Markdown -> DOCX pro Anfrage ausgewertet, ist standardmäßig deaktiviert und wird mit `[non_embed_links].auto_link_bare_url` in `configs/link.toml` aktiviert
- Markdown -> XLSX erzeugt keine DOCX-Hyperlink-Platzhalter und behält die ursprüngliche Link-Syntax bei

## 📖 Detaillierte Bedienungsanleitung

### Word zu Markdown

1.  Ziehen Sie die `.docx`-Datei in das Programmfenster.
2. Wählen Sie die Exportoptionen und klicken Sie auf **Markdown Export**.
3.  Generiert eine `.md`-Datei mit YAML-Metadaten.

**Unterstützte Formate**:
-   `.docx` - Standard Word-Dokument.
-   `.doc` - Automatisch in DOCX zur Verarbeitung konvertiert.
-   `.wps` - WPS-Dokument automatisch konvertiert.

**Exportoptionen**:

| Option | Beschreibung |
| :--- | :--- |
| **Bilder extrahieren** | Wenn aktiviert, werden Bilder im Dokument in den Ausgabeordner extrahiert und Bildlinks in die MD-Datei eingefügt. |
| **OCR** | Text in Bildern erkennen. Je nach OCR-Ausgabeeinstellung wird er in einer Markdown-Datei pro Bild oder in der Hauptdatei gespeichert. |
| **Optimierung** | Wenden Sie eine Optimierung für den gewählten Dokumenttyp an, etwa zum Extrahieren von Feldern aus chinesischen Amtsdokumenten. Ohne diese Option wird die Standardkonvertierung verwendet. |
| **Nr. entfernen** | Wenn aktiviert, werden Nummern vor Untertiteln entfernt (z.B. "一、", "（一）", "1." usw.) und in reinen Titeltext konvertiert. |
| **Nr. hinzufügen** | Wenn aktiviert, werden Nummern basierend auf Überschriftenebenen automatisch hinzugefügt (Nummerierungsschema kann in den Einstellungen konfiguriert werden). |

Überschriftsnummern aus mehrstufigen Word/WPS-Listen bleiben beim Export erhalten. Mit **Nr. entfernen** entfernen Sie die Nummern, ohne die Überschriftenebenen zu ändern.

### Markdown zu Word

1.  Bereiten Sie eine `.md`-Datei mit einem YAML-Header vor.
2.  Ziehen Sie sie in das Programmfenster und wählen Sie die entsprechende Word-Vorlage.
3. Wählen Sie DOCX als Zielformat und klicken Sie auf **Generieren DOCX**, um die Vorlage auszufüllen und das Dokument zu speichern.

**Konvertierungsoptionen**:

| Option | Beschreibung |
| :--- | :--- |
| **Nr. entfernen** | Wenn aktiviert, werden Nummern vor Untertiteln entfernt. |
| **Nr. hinzufügen** | Wenn aktiviert, werden Nummern basierend auf Überschriftenebenen automatisch hinzugefügt. |

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
    -   Neue verwaltete Stilnamen folgen der **Aufgabenausgabesprache** in den allgemeinen Einstellungen, nicht der Oberflächensprache. Ein UI-Sprachwechsel benennt bestehende Vorlagenstile nicht um.

**Vorschlag**: Nach dem Anpassen von Stilen in der Vorlage verwendet der Konverter automatisch Ihre Stile; wenn sie nicht in der Vorlage vorhanden sind, werden integrierte voreingestellte Stile verwendet.

### Tabellendateiverarbeitung

1. **Excel/CSV zu Markdown**: Fügen Sie eine `.xlsx`- oder `.csv`-Datei hinzu und klicken Sie auf **Markdown Export**.
2.  **Markdown zu Excel**: Markdown-Tabellen können nach XLSX exportiert werden. XLSX-Vorlagen unterstützen YAML-Felder, vertikale/horizontale Tabellenspalten-Platzhalter, Bildplatzhalter sowie verbundene und geschützte Zellen.

**Unterstützte Formate**:
-   `.xlsx` - Standard Excel-Dokument.
-   `.xls` - Automatisch in XLSX zur Verarbeitung konvertiert.
-   `.et` - WPS-Tabelle automatisch konvertiert.
-   `.csv` - CSV-Texttabelle.
-   `.tsv` - TSV-Tabelle (tabulatorgetrennt).


### Textkorrekturfunktion

Das Programm bietet vier anpassbare Korrekturregeln:

1.  **Zeichenpaarungsprüfung** - Erkennt, ob gepaarte Satzzeichen wie Klammern und Anführungszeichen übereinstimmen.
2.  **Symbolkorrektur** - Erkennt gemischte Verwendung von chinesischen und englischen Satzzeichen.
3.  **Tippfehlerprüfung** - Überprüft auf häufige Tippfehler basierend auf einem benutzerdefinierten Wörterbuch.
4.  **Sensitive Worterkennung** - Erkennt sensitive Wörter basierend auf einem benutzerdefinierten Wörterbuch.

**Benutzerdefinierte Wörterbücher**: Bearbeiten Sie Tippfehler- und sensitive Wörterbücher visuell in der "Einstellungen"-Oberfläche.

**Verwendung**:
1.  Ziehen Sie das zu prüfende Word-Dokument oder die Markdown-Datei in das Programm.
2.  Wählen Sie die erforderlichen Korrekturregeln.
3.  Klicken Sie auf die Schaltfläche "Textkorrektur".
4.  Korrekturergebnisse werden als Kommentare im Dokument angezeigt. Bei Markdown-Dateien wird ein JSON-Bericht ausgegeben.

Hinweis (JSON-Bericht für Markdown-Korrektur):
- Engine: `text_rules` + Markdown-Adapter `md_spell`
- Ausgabe: Der aktuelle CLI-Einstieg für Korrektur ist `validate`; verwenden Sie `--json` für die CLI-Hülle. `--report` ist ein optionaler Bericht-Dateipfad.

## 🛠️ Vorlagensystem

### Verwendung vorhandener Vorlagen

Verwalten Sie DOCX/XLSX-Vorlagen unter **Einstellungen → Vorlagen**: aktivieren, deaktivieren, sortieren und als Standard festlegen. Integrierte Vorlagen sind schreibgeschützt; eigene Vorlagen liegen im beschreibbaren Benutzerverzeichnis. Bei MSIX müssen Sie WindowsApps nicht ändern.

### Benutzerdefinierte Vorlagen

1. Kopieren und bearbeiten Sie eine integrierte Vorlage oder importieren Sie eine DOCX/XLSX-Datei.
2. Bearbeiten Sie Inhalt, Formatvorlagen und Platzhalter der eigenen Kopie in Word, Excel oder WPS; die Regeln stehen unten.
3. Speichern Sie die Datei und aktualisieren Sie die Liste in DocWen. Ein Neustart ist nicht erforderlich.

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

#### Excel-Vorlagenplatzhalter (Ziel zur Wiederherstellung früherer Funktionsgleichheit)

XLSX-Vorlagen unterstützen YAML-Feldplatzhalter, vertikale `{{↓Feld}}`- und horizontale `{{→Feld}}`-Tabellenspalten-Platzhalter, Bildplatzhalter sowie verbundene und geschützte Zellen.

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

**Behandlung verbundener Zellen**:

- Markdown -> Excel behält bestehende merged ranges der Vorlage bei.
- Für bekannte spaltenorientierte Vorlagenbereiche aus zusammenhängenden `{{↓Feldname}}`-Platzhaltern kann das Programm rechteckige Zusammenführungen aus expliziten `<` / `^`-Markern in Markdown-Tabellen wiederherstellen.
- Nur Zellen, deren getrimmter Inhalt exakt `<` oder `^` ist, nehmen an der Merge-Erkennung teil; `\<` und `\^` bleiben Literaltext.
- Ungültige Rechtecke oder Konflikte mit vorhandenen merged ranges der Vorlage werden mit Warnung zu normalem Text herabgestuft, statt die Vorlagenstruktur zwangsweise zu überschreiben.

**Zusammenführung mehrerer Tabellendaten**: Wenn es mehrere Tabellen in Markdown gibt, die denselben Kopfnamen verwenden, werden die Daten der Reihe nach zusammengeführt und nacheinander gefüllt.

## 🔌 Obsidian-Plugin

Ein begleitendes Obsidian-Plugin wird separat veröffentlicht und arbeitet mit dem Konverter zusammen:

### Kernfunktionen

-   **🚀 Ein-Klick-Start** - Seitenleistensymbol zum schnellen Starten des Konverters.
-   **📂 Automatische Übergabe** - Übergibt automatisch den aktuell geöffneten Dateipfad.
-   **🔄 Geöffnetes Fenster weiterverwenden** - Öffnet eine weitere Datei im laufenden Konverter, ohne ihn neu zu starten.

### Bedienung

DocWen Assistant 3.3 unterstützt Windows- und Linux-Desktop-Hosts. Microsoft-Store-Installationen unter Windows können automatisch erkannt werden; unter Linux wird ein kompatibles entpacktes Paket manuell ausgewählt. macOS wird vom Assistant derzeit nicht unterstützt.

1.  **Erster Klick** → Konverter starten und die aktuelle Datei öffnen.
2.  **Erneuter Klick mit geöffneter Datei** → Die bisherige Eingabedatei durch diese Datei ersetzen (Einzeldateimodus).
3.  **Erneuter Klick ohne geöffnete Datei** → Das Konverterfenster in den Vordergrund holen.

### Installation

Installieren Sie zuerst eine kompatible DocWen-Version und anschließend [DocWen Assistant](https://github.com/ZHYX91/obsidian-docwen-assistant). Assistant 3.3 erfordert DocWen 0.17.0 oder neuer; das Plugin installiert DocWen nicht mit. Prüfen Sie die Versionshinweise und laden Sie die passenden veröffentlichten Installationsdateien herunter.

## 🔌 OpenClaw (Plugin + Skill)

[docwen-openclaw](https://github.com/ZHYX91/docwen-openclaw) ermöglicht OpenClaw, mit DocWen lokale Dateien zu konvertieren und zu prüfen. Installieren Sie DocWen separat und wählen Sie eine Plugin-Version, deren Versionshinweise Ihre DocWen-Version als kompatibel nennen. Laden Sie die Dateien von der jeweiligen Release-Seite herunter; eine Versionsnummer im Quellcode bedeutet nicht, dass diese Version bereits zum Download verfügbar ist.

Für die Entwicklung von Integrationen mit Plugin 3.0 siehe [Machine Protocol v2 und Artifact Bundle v3](../specs/machine-protocol-v2.md).

## ❓ Häufig gestellte Fragen (FAQ)

### Was tun, wenn die Konvertierung fehlschlägt?

-   Prüfen Sie, ob die Datei von einem anderen Programm belegt ist.
-   Bestätigen Sie, dass das Dateiformat korrekt ist.
-   Prüfen Sie in den Einstellungen den Eintrag "Aktueller tatsächlicher Log-Dateipfad" oder sehen Sie im systemweiten Benutzer-Logverzeichnis nach den Fehlerprotokollen nach; wenn die Paketprüfung `DOCWEN_LOG_DIR` verwendet, prüfen Sie stattdessen das überschreibende Verzeichnis.

### Vorlage wird nicht angezeigt?

- Prüfen Sie unter **Einstellungen → Vorlagen**, ob die Vorlage aktiviert ist und zum Ausgabeformat passt.
- Prüfen Sie, ob sich die DOCX/XLSX-Datei öffnen lässt.
- Aktualisieren Sie die Liste nach externen Änderungen; neue Dateien können Sie importieren.

### Korrekturfunktion funktioniert nicht?

-   Bestätigen Sie, dass das Dokument im .docx- oder .md-Format vorliegt.
-   Prüfen Sie, ob das Dokument bearbeitbaren Text enthält.
-   Bestätigen Sie, dass Korrekturregeln in den Einstellungen aktiviert sind.

### Ausgabeformat nicht wie erwartet?

- Öffnen Sie **Einstellungen → Vorlagen**. Kopieren Sie integrierte Vorlagen vor dem Bearbeiten; eigene Vorlagen können Sie direkt bearbeiten.
- Passen Sie die Formatvorlagen in Word, Excel oder WPS an, speichern Sie und aktualisieren Sie die Liste.
- Wählen Sie die eigene Vorlage vor der nächsten Ausgabe aus.

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

DocWen verarbeitet Dateien auf Ihrem Computer und bietet weder Datei-Uploads noch Telemetrie, Online-Downloads oder einen Netzwerkdienst. Es beschränkt den Netzwerkzugriff seiner eigenen Python-Komponenten. Kann dieser Schutz nicht aktiviert werden, wird DocWen nicht weiter ausgeführt.

Separat gestartete Programme wie Word/Office, WPS und LibreOffice, einschließlich ihrer Hilfsprozesse, folgen ihren eigenen und den systemweiten Netzwerkeinstellungen. Sie fallen nicht unter diese Beschränkung. Prüfen Sie bei sensiblen Dateien daher auch deren Einstellungen.

Die [Architekturbeschreibung](../architecture.md) erklärt die Umsetzung und Grenzen des Schutzes. Bei Start- oder Konvertierungsfehlern hilft die [Fehlerbehebung](../maintenance/troubleshooting.md).

## 📜 Lizenz

Dieses Projekt ist unter der **GNU Affero General Public License v3.0 (AGPL-3.0)** lizenziert.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Dieses Projekt verwendet PyMuPDF (lizenziert unter AGPL-3.0), daher ist das gesamte Projekt ebenfalls unter AGPL-3.0 lizenziert.
- Die aktuelle GUI kann auf unterstützten Host-Pfaden `PySide6-Fluent-Widgets` (QFluentWidgets) verwenden; diese Abhängigkeit folgt einem `GPLv3 / kommerziell`-Doppellizenzmodell, während dieses Repository weiter unter AGPL verteilt wird.
-   Sie dürfen diese Software frei verwenden, ändern und verbreiten.
-   Wenn Sie diese Software ändern und Dienste über ein Netzwerk anbieten, müssen Sie den Benutzern den geänderten Quellcode zur Verfügung stellen.
-   Detaillierte Lizenzinformationen finden Sie in der Datei [LICENSE](../../LICENSE).
- Hinweise zu Drittanbieter-Komponenten finden Sie in [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt); eine Zusammenfassung der Weitergabe steht in [NOTICE.txt](../../NOTICE.txt).

### Kontakt

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Autor kontaktieren**: zhengyx91@hotmail.com

---

**Autor**: ZhengYX
