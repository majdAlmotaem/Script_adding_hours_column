from docx import Document
import os
from docx.shared import Inches
import glob

def add_hours_column(doc_path, output_path):
    # Dokument öffnen
    doc = Document(doc_path)
    
    # Stelle sicher, dass der Ausgabeordner existiert
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)

    # Finde die Tätigkeitstabelle (typischerweise die zweite Tabelle im Dokument)
    target_table = doc.tables[1]  # Index 1 für die zweite Tabelle

    # Neue Spalte zur Tätigkeitstabelle hinzufügen
    new_column = target_table.add_column(Inches(1))
    
    # Überschrift "Std"
    target_table.rows[0].cells[-1].text = "Std"

    # Füge Stunden hinzu
    for row in target_table.rows[1:]:
        # Prüfe den gesamten Zeileninhalt auf Urlaub/Feiertag
        row_text = ' '.join(cell.text.lower() for cell in row.cells).strip()
        if any(keyword in row_text for keyword in ['urlaub', 'feiertag']):
            row.cells[-1].text = ""
        else:
            row.cells[-1].text = "8"
    
    # Speichern
    doc.save(output_path)

# Basisverzeichnis
base_dir = "D:/Majd/IBB/Berichtsheft/Christoph_Backhaus_IT"
output_dir = os.path.join(base_dir, "bearbeitet")

# Erstelle Output-Verzeichnis falls nicht vorhanden
os.makedirs(output_dir, exist_ok=True)

# Finde alle .docx Dateien im Verzeichnis
docx_files = glob.glob(os.path.join(base_dir, "*.docx"))

# Verarbeite jede Datei
for old_name in docx_files:
    # Erstelle den neuen Dateinamen
    file_name = os.path.basename(old_name)
    new_name = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_Updated.docx")
    
    # Verarbeite die Datei
    add_hours_column(old_name, new_name)
    print(f"Verarbeitet: {file_name} -> {os.path.basename(new_name)}")

print(f"Alle Dateien wurden erfolgreich verarbeitet und im Ordner 'bearbeitet' gespeichert.")
