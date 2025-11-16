#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime, timedelta
import os
import sys

class GitHubMacConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Time Converter UTC+3")
        self.root.geometry("650x550")
        
        self.create_ui()
        
    def create_ui(self):
        """Создает кроссплатформенный UI"""
        # Header
        header = tk.Frame(self.root, bg="#2E86AB", height=80)
        header.pack(fill=tk.X, side=tk.TOP)
        
        title = tk.Label(header, text="🕒 Time Converter UTC+3", 
                        font=("Arial", 16, "bold"), bg="#2E86AB", fg="white")
        title.pack(pady=20)
        
        # Main content
        main = tk.Frame(self.root, padx=20, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        # File selection
        file_frame = tk.LabelFrame(main, text="Select Excel File", padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.file_var = tk.StringVar()
        file_entry = tk.Entry(file_frame, textvariable=self.file_var, width=50, state='readonly')
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(file_frame, text="Browse", command=self.browse_file, width=10)
        browse_btn.pack(side=tk.RIGHT)
        
        # Convert button
        self.convert_btn = tk.Button(main, text="Convert to UTC+3", 
                                   command=self.convert, state=tk.DISABLED,
                                   bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                   width=15, height=1)
        self.convert_btn.pack(pady=10)
        
        # Results area
        result_frame = tk.LabelFrame(main, text="Conversion Results", padx=10, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # Text area with scrollbar
        self.text_area = scrolledtext.ScrolledText(result_frame, height=15, width=70,
                                                 font=("Courier New", 9))
        self.text_area.pack(fill=tk.BOTH, expand=True)
        
        # Control buttons
        btn_frame = tk.Frame(result_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.copy_btn = tk.Button(btn_frame, text="Copy Results", 
                                command=self.copy_results, state=tk.DISABLED)
        self.copy_btn.pack(side=tk.LEFT)
        
        clear_btn = tk.Button(btn_frame, text="Clear", command=self.clear_results)
        clear_btn.pack(side=tk.RIGHT)
        
        # Status bar
        self.status = tk.StringVar(value="Ready to convert")
        status_bar = tk.Label(main, textvariable=self.status, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # Track file changes
        self.file_var.trace('w', self.on_file_change)
        
    def on_file_change(self, *args):
        path = self.file_var.get()
        if path and os.path.exists(path):
            self.convert_btn.config(state=tk.NORMAL, bg="#4CAF50")
            self.status.set("File selected - ready to convert")
        else:
            self.convert_btn.config(state=tk.DISABLED, bg="#CCCCCC")
            self.status.set("Please select an Excel file")
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_var.set(filename)
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, f"Selected: {os.path.basename(filename)}\n\n")
    
    def convert(self):
        input_file = self.file_var.get()
        if not input_file:
            return
            
        self.convert_btn.config(state=tk.DISABLED)
        self.status.set("Converting...")
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, "Converting time entries...\n")
        self.root.update()
        
        try:
            # Read and convert
            df = pd.read_excel(input_file, sheet_name='Meeting attendees')
            
            def convert_time(t):
                try:
                    dt = datetime.strptime(str(t), '%d-%m-%Y %H:%M')
                    return (dt + timedelta(hours=3)).strftime('%d-%m-%Y %H:%M')
                except:
                    return f"ERROR: {t}"
            
            df['Entry time'] = [convert_time(t) for t in df['Entry time']]
            
            # Save result
            base, ext = os.path.splitext(input_file)
            output_file = f"{base}_UTC+3{ext}"
            df.to_excel(output_file, index=False)
            
            # Display results
            result = self.format_results(df, input_file, output_file)
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, result)
            
            # Enable copy button
            self.copy_btn.config(state=tk.NORMAL)
            self.status.set(f"Conversion complete! {len(df)} records processed")
            
            messagebox.showinfo("Success", f"File saved as:\n{os.path.basename(output_file)}")
            
        except Exception as e:
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, f"❌ Error: {str(e)}")
            self.status.set("Conversion failed")
            messagebox.showerror("Error", str(e))
        finally:
            self.convert_btn.config(state=tk.NORMAL)
    
    def format_results(self, df, input_file, output_file):
        """Форматирует результаты для отображения"""
        result = "=" * 60 + "\n"
        result += "✅ CONVERSION SUCCESSFUL\n"
        result += "=" * 60 + "\n"
        result += f"Input:  {os.path.basename(input_file)}\n"
        result += f"Output: {os.path.basename(output_file)}\n\n"
        result += "CONVERTED DATA:\n"
        result += "-" * 60 + "\n"
        result += f"{'Name':<20} | {'Login':<12} | {'Entry Time'}\n"
        result += "-" * 60 + "\n"
        
        for _, row in df.iterrows():
            name = str(row['Name'])[:19]
            login = str(row['Login'])
            time_str = str(row['Entry time'])
            result += f"{name:<20} | {login:<12} | {time_str}\n"
        
        result += "-" * 60 + "\n"
        result += f"Total records: {len(df)}\n"
        
        return result
    
    def copy_results(self):
        text = self.text_area.get(1.0, tk.END)
        if text.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status.set("Results copied to clipboard")
    
    def clear_results(self):
        self.text_area.delete(1.0, tk.END)
        self.copy_btn.config(state=tk.DISABLED)
        self.status.set("Ready to convert")

def main():
    app = GitHubMacConverter()
    app.root.mainloop()

if __name__ == "__main__":
    main()
