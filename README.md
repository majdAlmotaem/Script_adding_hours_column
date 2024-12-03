# Stunden-Spalten-Generator für Word-Dokumente

Dieses Python-Skript automatisiert das Hinzufügen einer Stundenspalte zu Word-Dokumenten mit Tätigkeitstabellen. Es ist besonders nützlich für Berichtshefte oder ähnliche Dokumente, die eine Zeiterfassung benötigen.

## Funktionen

- Fügt automatisch eine "Std" (Stunden) Spalte zu Word-Dokumenten hinzu
- Trägt 8 Stunden für normale Arbeitstage ein
- Lässt die Stunden für Urlaub und Feiertage leer
- Verarbeitet mehrere Dokumente auf einmal
- Erstellt automatisch Sicherungskopien mit "_Bearbeitet" Suffix

## Installation

1. Stellen Sie sicher, dass Python auf Ihrem System installiert ist
2. Installieren Sie die benötigten Abhängigkeiten:
   ```bash
   pip install -r requirements.txt
   ```

## Verwendung

### Wichtig - So geben Sie die Pfade ein:

1. Öffnen Sie die Eingabeaufforderung (CMD)
2. Navigieren Sie zum Verzeichnis mit diesem Skript
3. Führen Sie einen der folgenden Befehle aus:

#### Option 1: Nur Eingabepfad
Das Skript erstellt automatisch einen 'bearbeitet' Unterordner für die Ausgabe:
```bash
python adding_hours_column.py "HIER_IHR_EINGABEPFAD"
```

Beispiel:
```bash
python adding_hours_column.py "C:/Users/IhrName/Dokumente/Berichte"
```

#### Option 2: Eingabe- UND Ausgabepfad
Wenn Sie den Ausgabeordner selbst bestimmen möchten:
```bash
python adding_hours_column.py "HIER_IHR_EINGABEPFAD" --output-dir "HIER_IHR_AUSGABEPFAD"
```

Beispiel:
```bash
python adding_hours_column.py "C:/Users/IhrName/Dokumente/Berichte" --output-dir "C:/Users/IhrName/Dokumente/Berichte/Bearbeitet"
```

### Wichtige Hinweise:

- Pfade müssen in Anführungszeichen stehen
- Verwenden Sie Schrägstriche (/) oder doppelte Backslashes (\\\\) in Pfaden
- Der Eingabepfad muss auf einen existierenden Ordner zeigen
- Der Ausgabeordner wird automatisch erstellt, falls er nicht existiert
- Originaldokumente bleiben unverändert

## Hilfe anzeigen

Um alle verfügbaren Optionen anzuzeigen:
```bash
python adding_hours_column.py --help
```
