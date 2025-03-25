import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
import datetime
import json
import os
from pathlib import Path
import sys

class PDFExcelProcessor:
    def __init__(self):
        self.config_file = 'config.json'
        self.excel_path = self.load_excel_path()
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("PDF to Excel Processor")
        self.root.geometry("600x400")

        # Create and configure drag & drop area
        self.drop_area = tk.Label(
            self.root,
            text="Drag and Drop PDF File Here",
            width=40,
            height=10,
            relief="solid"
        )
        self.drop_area.pack(pady=20)

        # Enable drag and drop
        self.drop_area.bind('<Drop>', self.handle_drop)
        self.drop_area.bind('<Enter>', self.handle_enter)

        # Create Excel file selection button
        self.excel_button = tk.Button(
            self.root,
            text="Select Excel File",
            command=self.select_excel_file
        )
        self.excel_button.pack(pady=10)

        # Display current Excel file path
        self.excel_label = tk.Label(
            self.root,
            text=f"Current Excel file: {self.excel_path or 'None'}"
        )
        self.excel_label.pack(pady=10)

        # Make window accept drag and drop
        self.root.drop_target_register(tk.DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def load_excel_path(self):
        """Load saved Excel path from config file"""
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
                return config.get('excel_path')
        except FileNotFoundError:
            return None

    def save_excel_path(self, path):
        """Save Excel path to config file"""
        with open(self.config_file, 'w') as f:
            json.dump({'excel_path': path}, f)

    def select_excel_file(self):
        """Handle Excel file selection"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.excel_path = file_path
            self.save_excel_path(file_path)
            self.excel_label.config(text=f"Current Excel file: {file_path}")

    def handle_enter(self, event):
        """Handle drag enter event"""
        event.widget.focus_force()
        return tk.NONE

    def handle_drop(self, event):
        """Handle file drop event"""
        file_path = event.data
        if file_path.lower().endswith('.pdf'):
            self.process_pdf(file_path)
        else:
            messagebox.showerror("Error", "Please drop a PDF file")

    def process_pdf(self, pdf_path):
        """Process the PDF file and update Excel"""
        if not self.excel_path:
            messagebox.showerror("Error", "Please select an Excel file first")
            return

        try:
            # Read the main sheet with student names
            excel_df = pd.read_excel(self.excel_path, sheet_name='Main')
            student_names = excel_df['Student'].tolist()

            # Extract data from PDF
            pdf_data = []
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table[1:]:  # Skip header row
                            if row[0] in student_names:  # Check if student name matches
                                pdf_data.append(row)

            # Create new sheet with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            columns = ['Student', 'Section', 'Title', 'Category', 
                      'Assignment', 'Due Date', 'Status']
            
            new_df = pd.DataFrame(pdf_data, columns=columns)
            
            # Read existing Excel file
            with pd.ExcelWriter(self.excel_path, mode='a', engine='openpyxl') as writer:
                new_df.to_excel(writer, sheet_name=timestamp, index=False)

            messagebox.showinfo("Success", 
                              f"Data processed successfully!\nNew sheet '{timestamp}' created.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def run(self):
        """Start the application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = PDFExcelProcessor()
    app.run() 