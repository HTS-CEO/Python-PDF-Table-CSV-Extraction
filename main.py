import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import csv
import fitz  
from threading import Thread

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Table Extractor")
        self.root.geometry("700x500")
        
        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()

        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        title_label = ttk.Label(main_frame, text="PDF Table Extractor", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=10, sticky=tk.W)
        
        ttk.Label(main_frame, text="Input Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Entry(input_frame, textvariable=self.input_dir, width=60).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(row=0, column=1)
        
        # Output file selection
        ttk.Label(main_frame, text="Output CSV File:").grid(row=3, column=0, sticky=tk.W, pady=5)
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Entry(output_frame, textvariable=self.output_file, width=60).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(output_frame, text="Browse", command=self.browse_output).grid(row=0, column=1)
        
        # Progress bar
        ttk.Label(main_frame, text="Progress:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status = ttk.Label(main_frame, text="Ready to extract tables from PDFs")
        self.status.grid(row=7, column=0, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, pady=10)
        ttk.Button(button_frame, text="Extract Tables", command=self.start_extraction).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Clear", command=self.clear_fields).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).grid(row=0, column=2, padx=5)
        
        # Results area
        ttk.Label(main_frame, text="Processing Log:").grid(row=9, column=0, sticky=tk.W, pady=5)
        
        # Create a frame for the text widget and scrollbar
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=10, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Add a text widget with a scrollbar
        self.log_text = tk.Text(log_frame, height=12, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(10, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def browse_input(self):
        directory = filedialog.askdirectory(title="Select PDF Directory")
        if directory:
            self.input_dir.set(directory)
            
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save CSV As",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            
    def clear_fields(self):
        self.input_dir.set("")
        self.output_file.set("")
        self.log_text.delete(1.0, tk.END)
        self.status.config(text="Ready to extract tables from PDFs")
        self.progress['value'] = 0
        
    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_extraction(self):
        if not self.input_dir.get():
            messagebox.showerror("Error", "Please select input directory")
            return
            
        if not self.output_file.get():
            messagebox.showerror("Error", "Please select output file")
            return
            
        # Disable buttons during processing
        self.toggle_buttons(False)
        self.progress['value'] = 0
        self.status.config(text="Processing...")
        
        # Run extraction in a separate thread
        thread = Thread(target=self.extract_tables)
        thread.daemon = True
        thread.start()
        
    def toggle_buttons(self, state):
        for child in self.root.winfo_children():
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL if state else tk.DISABLED)
                
    def extract_tables(self):
        try:
            input_dir = self.input_dir.get()
            output_csv = self.output_file.get()
            
            all_rows = []
            pdf_files = []
            
            # Get all PDF files
            for filename in os.listdir(input_dir):
                if filename.lower().endswith('.pdf'):
                    pdf_files.append(filename)
            
            if not pdf_files:
                self.log_message("No PDF files found in the directory")
                self.status.config(text="No PDF files found")
                self.toggle_buttons(True)
                return
                
            self.log_message(f"Found {len(pdf_files)} PDF files")
            self.log_message("Starting PDF extraction...")
            
            processed_files = 0
            for filename in sorted(pdf_files):
                pdf_path = os.path.join(input_dir, filename)
                self.log_message(f"Processing: {filename}")
                
                try:
                    # Open the PDF
                    doc = fitz.open(pdf_path)
                    
                    # Extract text from each page
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        text = page.get_text()
                        
                        # Split text into lines and then into columns
                        lines = text.split('\n')
                        for line in lines:
                            # Simple tabular data detection - split by multiple spaces
                            row = [cell.strip() for cell in line.split('  ') if cell.strip()]
                            if len(row) > 1:  # Consider it a table row if it has multiple columns
                                all_rows.append(row)
                    
                    processed_files += 1
                    self.log_message(f"✓ Successfully processed {filename}")
                    
                    # Update progress
                    self.progress['value'] = (processed_files / len(pdf_files)) * 100
                    self.root.update_idletasks()
                    
                except Exception as e:
                    self.log_message(f"✗ Error processing {filename}: {str(e)}")
            
            if all_rows:
                # Write to CSV
                with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(all_rows)
                
                self.log_message(f"CSV saved to: {output_csv}")
                self.log_message(f"Processed {processed_files} PDF files with {len(all_rows)} total rows")
                self.status.config(text=f"Completed: {processed_files} files processed, {len(all_rows)} rows extracted")
            else:
                self.log_message("No table data found in the PDF files")
                self.status.config(text="No table data found")
                
        except Exception as e:
            self.log_message(f"Unexpected error: {str(e)}")
            self.status.config(text="Error occurred")
            
        finally:
            self.toggle_buttons(True)

def main():
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
