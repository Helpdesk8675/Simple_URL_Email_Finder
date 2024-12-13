# pip install python-magic-bin PyPDF2 python-docx

import tkinter as tk
from tkinter import filedialog, ttk
import os
import re
import datetime
import csv
import PyPDF2
import docx
import magic  # for file type detection
import threading
from pathlib import Path

class URLEmailExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("URL and Email Extractor")
        self.root.geometry("600x200")
        
        # Input folder selection
        self.input_frame = ttk.Frame(root)
        self.input_frame.pack(fill='x', padx=5, pady=5)
        
        self.input_label = ttk.Label(self.input_frame, text="Input Folder:")
        self.input_label.pack(side='left')
        
        self.input_path = tk.StringVar()
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path, width=50)
        self.input_entry.pack(side='left', padx=5)
        
        self.input_button = ttk.Button(self.input_frame, text="Browse", command=self.browse_input)
        self.input_button.pack(side='left')
        
        # Output folder selection
        self.output_frame = ttk.Frame(root)
        self.output_frame.pack(fill='x', padx=5, pady=5)
        
        self.output_label = ttk.Label(self.output_frame, text="Output Folder:")
        self.output_label.pack(side='left')
        
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path, width=50)
        self.output_entry.pack(side='left', padx=5)
        
        self.output_button = ttk.Button(self.output_frame, text="Browse", command=self.browse_output)
        self.output_button.pack(side='left')
        
        # Progress bar
        self.progress_frame = ttk.Frame(root)
        self.progress_frame.pack(fill='x', padx=5, pady=5)
        
        self.progress = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress.pack(fill='x')
        
        # Start button
        self.start_button = ttk.Button(root, text="Start Extraction", command=self.start_extraction)
        self.start_button.pack(pady=10)
        
        # Status label
        self.status_label = ttk.Label(root, text="")
        self.status_label.pack()

    def browse_input(self):
        folder = filedialog.askdirectory()
        self.input_path.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        self.output_path.set(folder)

    def extract_from_pdf(self, file_path):
        urls = set()
        emails = set()
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    self.find_urls_and_emails(text, urls, emails)
        except:
            pass
        return urls, emails

    def extract_from_docx(self, file_path):
        urls = set()
        emails = set()
        try:
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                self.find_urls_and_emails(para.text, urls, emails)
        except:
            pass
        return urls, emails

    def find_urls_and_emails(self, text, urls, emails):
        # URL regex pattern
        url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
        # Email regex pattern
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        
        urls.update(re.findall(url_pattern, text))
        emails.update(re.findall(email_pattern, text))

    def process_file(self, file_path):
        urls = set()
        emails = set()
        
        mime = magic.Magic(mime=True)
        file_type = mime.from_file(file_path)
        
        if 'pdf' in file_type.lower():
            urls, emails = self.extract_from_pdf(file_path)
        elif 'officedocument' in file_type.lower():
            urls, emails = self.extract_from_docx(file_path)
        else:
            # Try to read as text
            try:
                with open(file_path, 'r', errors='ignore') as file:
                    text = file.read()
                    self.find_urls_and_emails(text, urls, emails)
            except:
                pass
                
        return urls, emails

    def start_extraction(self):
        if not self.input_path.get() or not self.output_path.get():
            self.status_label.config(text="Please select both input and output folders")
            return
            
        self.start_button.config(state='disabled')
        self.progress.start()
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        output_file = os.path.join(self.output_path.get(), f"{timestamp}_EmailURLsearch.csv")
        
        results = []
        
        for root, _, files in os.walk(self.input_path.get()):
            for file in files:
                file_path = os.path.join(root, file)
                urls, emails = self.process_file(file_path)
                
                # Add results to list
                for url in urls:
                    results.append([file_path, "URL", url])
                for email in emails:
                    results.append([file_path, "Email", email])
        
        # Write results to CSV
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["File", "Type", "Value"])
            writer.writerows(results)
        
        self.progress.stop()
        self.start_button.config(state='normal')
        self.status_label.config(text=f"Extraction complete. Results saved to {output_file}")

def main():
    root = tk.Tk()
    app = URLEmailExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
