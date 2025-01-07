from docx import Document
import os
from docx.shared import Inches
import glob
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def get_unprocessed_files(input_files, output_dir):
    unprocessed = []
    for input_file in input_files:
        file_name = os.path.basename(input_file)
        output_file = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_Processed.docx")
        if not os.path.exists(output_file):
            unprocessed.append(input_file)
    return unprocessed

def add_hours_column(doc_path, output_path):
    doc = Document(doc_path)
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)
    target_table = doc.tables[1]
    new_column = target_table.add_column(Inches(1))
    target_table.rows[0].cells[-1].text = "Hours"
    
    for row in target_table.rows[1:]:
        row_text = ' '.join(cell.text.lower() for cell in row.cells).strip()
        if any(keyword in row_text for keyword in ['vacation', 'holiday']):
            row.cells[-1].text = ""
        else:
            row.cells[-1].text = "8"
    
    doc.save(output_path)

class HoursColumnManagerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Hours Column Manager")
        self.root.geometry("600x400")
        
        # Add icon using absolute path
        icon = tk.PhotoImage(file='hours column.png')
        self.root.iconphoto(True, icon)
        
        # Create Menu Bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Create Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=self.show_help)
        
        # Dark Mode Theme
        self.root.configure(bg='#2b2b2b')
        style = ttk.Style()
        style.theme_use('clam')  
        
        # Configure dark mode colors
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
        style.configure('TButton', background='#404040', foreground='#ffffff')
        style.configure('TEntry', fieldbackground='#404040', foreground='#ffffff')
        
        # Center window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 400) // 2
        self.root.geometry(f"600x400+{x}+{y}")
        
        # Main container with center alignment
        self.main_frame = ttk.Frame(self.root, padding="20", style='TFrame')
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Input folder
        ttk.Label(self.main_frame, text="Input folder:", style='TLabel').grid(row=0, column=0, sticky="e", pady=5)
        self.input_path = tk.StringVar()
        self.input_entry = ttk.Entry(self.main_frame, textvariable=self.input_path, width=50, style='TEntry')
        self.input_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_input, style='TButton').grid(row=0, column=2, pady=5, padx=5)
        
        # Output folder
        ttk.Label(self.main_frame, text="Output folder:", style='TLabel').grid(row=1, column=0, sticky="e", pady=5)
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.main_frame, textvariable=self.output_path, width=50, style='TEntry')
        self.output_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_output, style='TButton').grid(row=1, column=2, pady=5, padx=5)
        
        # Start Button - centered
        start_button = ttk.Button(self.main_frame, text="Add Hours Column", command=self.process_files, style='TButton')
        start_button.grid(row=2, column=0, columnspan=3, pady=20)
        
        # Status Label - centered
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, wraplength=500, justify="center", style='TLabel')
        self.status_label.grid(row=3, column=0, columnspan=3, pady=5)

        # Configure grid columns to center content
        self.main_frame.grid_columnconfigure(1, weight=1)



    # Füge die show_help Methode direkt nach __init__ ein, auf gleicher Ebene wie andere Methoden (browse_input, process_files etc.)
    def show_help(self):
        help_text = """
How to use Hours Column Manager:

1. Select Input Folder
   - Click 'Browse' to choose the folder containing your .docx files
   - Files must be Word documents with tables

2. Select Output Folder (Optional)
   - Click 'Browse' to choose where to save processed files
   - If not selected, a 'processed' folder will be created automatically

3. Process Files
   - Click 'Add Hours Column' to start processing
   - New files will be saved with '_Processed' suffix
   - Hours column (8h) will be added automatically
   - Vacation/Holiday entries will be left empty

Note: Already processed files will be skipped automatically.
"""
        messagebox.showinfo("Help", help_text)



    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_path.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(folder)

    def process_files(self):
        input_dir = self.input_path.get()
        output_dir = self.output_path.get()

        if not input_dir:
            messagebox.showerror("Error", "Please select an input folder!")
            return

        if not output_dir:
            output_dir = os.path.join(input_dir, "processed")
            self.output_path.set(output_dir)

        docx_files = glob.glob(os.path.join(input_dir, "*.docx"))
        if not docx_files:
            messagebox.showinfo("Info", "No .docx files found in input folder!")
            return

        unprocessed_files = get_unprocessed_files(docx_files, output_dir)
        if not unprocessed_files:
            messagebox.showinfo("Info", "No new files to process!")
            return

        for old_name in unprocessed_files:
            file_name = os.path.basename(old_name)
            new_name = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}_Processed.docx")
            self.status_var.set(f"Processing: {file_name}")
            self.root.update()
            add_hours_column(old_name, new_name)

        messagebox.showinfo("Info", f"{len(unprocessed_files)} files have been successfully processed!")
        self.status_var.set("Ready")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = HoursColumnManagerGUI()
    app.run()
