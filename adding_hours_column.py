"""
Fügt eine Stundenspalte zu Word-Dokumenten mit Tätigkeitstabellen hinzu.
Weitere Informationen zur Verwendung finden Sie in der README.md
"""

from docx import Document
import os
from docx.shared import Inches
import glob
import argparse

def add_hours_column(doc_path, output_path):
    # Dokument öffnen
    doc = Document(doc_path)
    
    # Ausgabeordner erstellen
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)

    # Zweite Tabelle im Dokument bearbeiten
    target_table = doc.tables[1]
    new_column = target_table.add_column(Inches(1))
    target_table.rows[0].cells[-1].text = "Std"

    # Stunden eintragen (8 für Arbeitstage, leer für Urlaub/Feiertag)
    for row in target_table.rows[1:]:
        row_text = ' '.join(cell.text.lower() for cell in row.cells).strip()
        if any(keyword in row_text for keyword in ['urlaub', 'feiertag']):
            row.cells[-1].text = ""
        else:
            row.cells[-1].text = "8"
    
    doc.save(output_path)

def main():
    parser = argparse.ArgumentParser(
        description='Fügt eine Stundenspalte zu Word-Dokumenten hinzu.'
    )
    
    parser.add_argument('input_dir', help='Verzeichnis mit den Word-Dokumenten')
    parser.add_argument('--output-dir', help='Ausgabeverzeichnis (Optional)', default=None)
    
    args = parser.parse_args()
    base_dir = os.path.abspath(args.input_dir)
    output_dir = os.path.abspath(args.output_dir) if args.output_dir else os.path.join(base_dir, "bearbeitet")

    if not os.path.exists(base_dir):
        print(f"Fehler: Eingabeverzeichnis '{base_dir}' existiert nicht!")
        return

    os.makedirs(output_dir, exist_ok=True)
    docx_files = glob.glob(os.path.join(base_dir, "*.docx"))
    
    if not docx_files:
        print(f"Keine .docx Dateien in '{base_dir}' gefunden")
        return

    print(f"{len(docx_files)} Word-Dokumente gefunden...")
    
    for old_name in docx_files:
        file_name = os.path.basename(old_name)
        new_name = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_Bearbeitet.docx")
        
        print(f"Verarbeite: {file_name}")
        add_hours_column(old_name, new_name)
        print(f"Gespeichert als: {os.path.basename(new_name)}")

    print(f"\nAlle Dateien wurden erfolgreich in '{output_dir}' gespeichert.")

if __name__ == "__main__":
    main()
