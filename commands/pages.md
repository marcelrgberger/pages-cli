---
name: pages
description: "Apple Pages steuern — /pages new, /pages open <pfad>, /pages export pdf <pfad>, /pages templates, /pages info, /pages close, /pages status"
arguments:
  - name: action
    description: "Aktion: new [Vorlage], open <Pfad>, info, close, export <format> <pfad>, templates, status"
    required: false
---

# Pages CLI — Schnellzugriff

Du hast Zugriff auf die Apple Pages CLI (`cli-anything-pages`).

## Schritt 0: CLI sicherstellen

Prüfe ob `cli-anything-pages` verfügbar ist. Falls nicht, installiere es:

```bash
which cli-anything-pages || (cd ${CLAUDE_PLUGIN_ROOT}/agent-harness && python3 -m venv .venv && source .venv/bin/activate && pip install -e . && echo "cli-anything-pages installed")
```

Falls eine venv existiert, aktiviere sie:

```bash
test -f ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate && source ${CLAUDE_PLUGIN_ROOT}/agent-harness/.venv/bin/activate
```

## Deine Aufgabe

Parse das Argument `$ARGUMENTS` und führe die passende Aktion aus:

### Wenn kein Argument oder "new":
1. Erstelle ein neues leeres Dokument:
   ```bash
   cli-anything-pages --json document new
   ```
2. Informiere den User und warte auf Anweisungen was er damit machen möchte.

### Wenn "new <Vorlagenname>":
1. Erstelle ein Dokument aus der genannten Vorlage:
   ```bash
   cli-anything-pages --json document new --template "<Vorlagenname>"
   ```
2. Informiere den User und warte auf Anweisungen.

### Wenn "open <Pfad>":
1. Öffne das Dokument:
   ```bash
   cli-anything-pages --json document open "<Pfad>"
   ```
2. Hole Dokumentinfo:
   ```bash
   cli-anything-pages --json document info
   ```
3. Zeige dem User eine Zusammenfassung (Name, Seiten, Wörter) und warte auf Anweisungen.

### Wenn "info":
```bash
cli-anything-pages --json document info
cli-anything-pages --json text word-count
```
Zeige die Informationen übersichtlich an.

### Wenn "close":
```bash
cli-anything-pages document close
```

### Wenn "export pdf <Pfad>" oder "export word <Pfad>" etc.:
```bash
cli-anything-pages export <format> "<Pfad>"
```
Bestätige den Export mit Dateigröße.

### Wenn "templates":
```bash
cli-anything-pages --json template list
```
Zeige die Vorlagen gruppiert an (Berichte, Briefe, CVs, etc.).

### Wenn "status":
```bash
cli-anything-pages --json session status
```

## Nach dem Öffnen/Erstellen

Sobald ein Dokument offen ist, übersetze natürliche Sprache des Users in CLI-Befehle:

- Text schreiben → `cli-anything-pages text add "..."`
- Formatierung → `cli-anything-pages text set-font --name "..." --size N`
- Tabelle → `cli-anything-pages table add --rows N --cols N`
- Bild → `cli-anything-pages media add-image "Pfad"`
- Export → `cli-anything-pages export pdf/word/epub "Pfad"`
- Review → `cli-anything-pages --json text get` → Analysiere den Text und schlage Verbesserungen vor
- Formatierung prüfen → `cli-anything-pages --json document info` → Prüfe Konsistenz

Verwende IMMER `--json` bei Abfrage-Befehlen für strukturierte Ausgabe.
