import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from generate_cards import generate_cards
from convert_data import convert_data_to_format
import threading

class CardGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Laabharti Card Generator")
        self.root.geometry("600x500")  # Made taller to accommodate the additional logo section
        self.root.resizable(True, True)
        
        # Set up the main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Input CSV File", padding="10")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.csv_path = tk.StringVar()
        csv_entry = ttk.Entry(file_frame, textvariable=self.csv_path, width=50)
        csv_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W+tk.E)
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_csv)
        browse_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Custom logo section
        logo_frame = ttk.LabelFrame(main_frame, text="Custom Logo Image (Optional)", padding="10")
        logo_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.logo_path = tk.StringVar()
        logo_entry = ttk.Entry(logo_frame, textvariable=self.logo_path, width=50)
        logo_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W+tk.E)
        
        logo_btn = ttk.Button(logo_frame, text="Browse", command=self.browse_logo)
        logo_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Output file section
        output_frame = ttk.LabelFrame(main_frame, text="Output Word Document", padding="10")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.output_path = tk.StringVar(value="output_cards.docx")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W+tk.E)
        
        save_btn = ttk.Button(output_frame, text="Save As", command=self.save_as)
        save_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Layout options section
        layout_frame = ttk.LabelFrame(main_frame, text="Card Layout", padding="10")
        layout_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(layout_frame, text="Rows per page:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.rows = tk.IntVar(value=4)
        rows_spinbox = ttk.Spinbox(layout_frame, from_=1, to=10, textvariable=self.rows, width=5)
        rows_spinbox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(layout_frame, text="Columns per page:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.cols = tk.IntVar(value=2)
        cols_spinbox = ttk.Spinbox(layout_frame, from_=1, to=10, textvariable=self.cols, width=5)
        cols_spinbox.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Data conversion option
        self.convert_data = tk.BooleanVar(value=False)
        convert_check = ttk.Checkbutton(layout_frame, text="Auto-convert CSV format", variable=self.convert_data)
        convert_check.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # Status section
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(fill=tk.X, padx=5, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=200, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=5, pady=5)
        
        # Buttons section
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        generate_btn = ttk.Button(button_frame, text="Generate Cards", command=self.generate)
        generate_btn.pack(side=tk.RIGHT, padx=5)
        
        quit_btn = ttk.Button(button_frame, text="Quit", command=root.destroy)
        quit_btn.pack(side=tk.RIGHT, padx=5)
    
    def browse_csv(self):
        """Open file dialog to select CSV file"""
        filepath = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filepath:
            self.csv_path.set(filepath)
    
    def browse_logo(self):
        """Open file dialog to select logo image"""
        filepath = filedialog.askopenfilename(
            title="Select Logo Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        if filepath:
            self.logo_path.set(filepath)
    
    def save_as(self):
        """Open file dialog to select save location"""
        filepath = filedialog.asksaveasfilename(
            title="Save Word Document As",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")]
        )
        if filepath:
            self.output_path.set(filepath)
    
    def process_files(self):
        """Process the files in a separate thread"""
        try:
            csv_path = self.csv_path.get()
            output_path = self.output_path.get()
            logo_path = self.logo_path.get() if self.logo_path.get() else None
            rows = self.rows.get()
            cols = self.cols.get()
            
            if not csv_path:
                messagebox.showerror("Error", "Please select a CSV file.")
                return
            
            if not output_path:
                messagebox.showerror("Error", "Please specify an output file.")
                return
            
            if not os.path.exists(csv_path):
                messagebox.showerror("Error", f"CSV file not found: {csv_path}")
                return
            
            # Convert data if needed
            if self.convert_data.get():
                self.status_var.set("Converting CSV data format...")
                converted_csv = os.path.join(os.path.dirname(csv_path), 'converted_data.csv')
                if not convert_data_to_format(csv_path, converted_csv):
                    messagebox.showerror("Error", "Failed to convert CSV format.")
                    self.status_var.set("Ready")
                    self.progress.stop()
                    return
                csv_path = converted_csv
            
            # Generate cards
            self.status_var.set("Generating cards...")
            generate_cards(csv_path, output_path, rows, cols, logo_path)
            
            self.status_var.set("Cards generated successfully!")
            messagebox.showinfo("Success", f"Cards generated successfully and saved to {output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
        finally:
            self.progress.stop()
    
    def generate(self):
        """Start card generation process in a separate thread"""
        self.progress.start()
        threading.Thread(target=self.process_files, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = CardGeneratorApp(root)
    root.mainloop() 