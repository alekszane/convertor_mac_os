#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime, timedelta
import os
import json
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

# Файл для сохранения настроек
SETTINGS_FILE = os.path.expanduser("~/Library/Application Support/TimeConverter/settings.json")

def get_settings_dir():
    """Создает папку для настроек в Application Support"""
    settings_dir = os.path.dirname(SETTINGS_FILE)
    if not os.path.exists(settings_dir):
        os.makedirs(settings_dir)
    return settings_dir

def load_settings():
    """Загружает сохраненные настройки окна"""
    default_settings = {
        "window_width": 650,
        "window_height": 550,
        "window_x": None,
        "window_y": None
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                for key in default_settings:
                    if key not in settings:
                        settings[key] = default_settings[key]
                return settings
    except:
        pass
    return default_settings

def save_settings(settings):
    """Сохраняет настройки окна"""
    try:
        get_settings_dir()
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except:
        pass

def convert_time_value(time_str):
    """Конвертирует время в UTC+3"""
    try:
        if pd.isna(time_str) or str(time_str).strip() == '':
            return time_str
        original_time = datetime.strptime(str(time_str), '%d-%m-%Y %H:%M')
        utc3_time = original_time + timedelta(hours=3)
        return utc3_time.strftime('%d-%m-%Y %H:%M')
    except:
        return time_str

def convert_excel_file(input_file):
    """Конвертирует Excel файл"""
    try:
        name, ext = os.path.splitext(input_file)
        output_file = f"{name}_UTC+3{ext}"
        df = pd.read_excel(input_file, sheet_name='Meeting attendees')
        
        time_columns = ['Entry time', 'Exit time']
        converted_counts = {}
        
        for col in time_columns:
            if col in df.columns:
                original_count = df[col].notna().sum()
                df[col] = df[col].apply(convert_time_value)
                converted_counts[col] = original_count
        
        df.to_excel(output_file, sheet_name='Meeting attendees', index=False)
        return True, output_file, df, list(converted_counts.keys()), converted_counts
    except Exception as e:
        return False, str(e), None, [], {}

class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🕒 Time Converter UTC+3")
        
        # macOS стиль
        try:
            self.root.tk.call('tk', 'mac', 'style', 'use', 'system')
        except:
            pass
        
        # Загрузка настроек
        self.settings = load_settings()
        width = self.settings.get("window_width", 650)
        height = self.settings.get("window_height", 550)
        x = self.settings.get("window_x")
        y = self.settings.get("window_y")
        
        if x and y:
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.root.geometry(f"{width}x{height}")
            self.center_window()
        
        self.root.minsize(600, 500)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind("<Configure>", self.on_window_resize)
        self.root.configure(bg='#f0f0f0')
        
        self.setup_ui()
        self.full_data = ""
    
    def center_window(self):
        self.root.update_idletasks()
        w = self.settings.get("window_width", 650)
        h = self.settings.get("window_height", 550)
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.settings["window_x"] = x
        self.settings["window_y"] = y
    
    def on_window_resize(self, event):
        if event.widget == self.root:
            self.settings["window_width"] = self.root.winfo_width()
            self.settings["window_height"] = self.root.winfo_height()
            self.settings["window_x"] = self.root.winfo_x()
            self.settings["window_y"] = self.root.winfo_y()
            save_settings(self.settings)
    
    def on_closing(self):
        save_settings(self.settings)
        if messagebox.askokcancel("Quit", "Close application?"):
            self.root.destroy()
    
    def setup_ui(self):
        main = tk.Frame(self.root, bg='#f0f0f0', padx=20, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        tk.Label(main, text="🕒 TIME CONVERTER UTC+3", font=("Helvetica", 14, "bold"),
                fg="#2E86AB", bg='#f0f0f0').pack(pady=(0, 5))
        tk.Label(main, text="Converts Entry time and Exit time to UTC+3",
                font=("Helvetica", 9), fg="#666666", bg='#f0f0f0').pack(pady=(0, 15))
        
        # Выбор файла
        file_frame = tk.LabelFrame(main, text=" Select File ", font=("Helvetica", 10, "bold"),
                                   bg='#f0f0f0', padx=12, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 12))
        
        path_frame = tk.Frame(file_frame, bg='#f0f0f0')
        path_frame.pack(fill=tk.X)
        
        self.file_path = tk.StringVar()
        tk.Entry(path_frame, textvariable=self.file_path, font=("Monaco", 9),
                bg='white', relief=tk.FLAT).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Button(path_frame, text="📁 Browse", command=self.browse_file,
                 font=("Helvetica", 9), bg="#E8E8E8", relief=tk.FLAT, padx=12).pack(side=tk.RIGHT)
        
        # Кнопка конвертации
        self.convert_btn = tk.Button(main, text="🔄 Convert to UTC+3",
                                     command=self.start_conversion, font=("Helvetica", 11, "bold"),
                                     bg="#4CD964", fg="white", relief=tk.FLAT, padx=20, pady=6,
                                     state=tk.DISABLED)
        self.convert_btn.pack(pady=(10, 12))
        
        # Результаты
        result_frame = tk.LabelFrame(main, text=" Conversion Results ", font=("Helvetica", 10, "bold"),
                                     bg='#f0f0f0', padx=12, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # Кнопки управления
        controls = tk.Frame(result_frame, bg='#f0f0f0')
        controls.pack(fill=tk.X, pady=(0, 8))
        
        self.copy_btn = tk.Button(controls, text="📋 Copy All", command=self.copy_all,
                                  font=("Helvetica", 8), bg="#4CAF50", fg="white", relief=tk.FLAT,
                                  padx=10, pady=3, state=tk.DISABLED)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        self.copy_vis_btn = tk.Button(controls, text="📋 Copy Visible", command=self.copy_visible,
                                      font=("Helvetica", 8), bg="#E8E8E8", relief=tk.FLAT,
                                      padx=10, pady=3, state=tk.DISABLED)
        self.copy_vis_btn.pack(side=tk.LEFT)
        
        self.records_var = tk.StringVar()
        tk.Label(controls, textvariable=self.records_var, font=("Helvetica", 8, "bold"),
                fg="#2E86AB", bg='#f0f0f0').pack(side=tk.RIGHT)
        
        tk.Button(controls, text="🗑️ Clear", command=self.clear_results,
                 font=("Helvetica", 8), bg="#E8E8E8", relief=tk.FLAT, padx=10, pady=3).pack(side=tk.RIGHT, padx=(0, 5))
        
        # Текстовое поле с прокруткой
        text_container = tk.Frame(result_frame, bg='#f0f0f0')
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.result_text = tk.Text(text_container, wrap=tk.NONE, font=("Monaco", 9),
                                   bg="#fafafa", padx=10, pady=10, relief=tk.FLAT)
        
        v_scroll = tk.Scrollbar(text_container, orient=tk.VERTICAL, command=self.result_text.yview)
        h_scroll = tk.Scrollbar(text_container, orient=tk.HORIZONTAL, command=self.result_text.xview)
        self.result_text.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Статус
        self.status_var = tk.StringVar(value="Ready - select an Excel file")
        tk.Label(main, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W,
                font=("Helvetica", 8), bg="#e8e8e8", padx=8, pady=3).pack(fill=tk.X, pady=(8, 0))
        
        self.file_path.trace_add('write', self.check_file)
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_path.set(filename)
            self.clear_results()
    
    def check_file(self, *args):
        path = self.file_path.get()
        if path and os.path.exists(path):
            self.convert_btn.config(state=tk.NORMAL, bg="#4CD964")
            self.status_var.set("File selected - ready to convert")
        else:
            self.convert_btn.config(state=tk.DISABLED, bg="#CCCCCC")
            self.status_var.set("Select an Excel file")
    
    def start_conversion(self):
        self.convert_btn.config(state=tk.DISABLED, bg="#CCCCCC")
        threading.Thread(target=self.convert_file, daemon=True).start()
    
    def convert_file(self):
        input_file = self.file_path.get()
        self.status_var.set("Converting...")
        
        try:
            success, result, df, cols, counts = convert_excel_file(input_file)
            if success:
                self.root.after(0, lambda: self.show_success(input_file, result, df, cols, counts))
            else:
                self.root.after(0, lambda: self.show_error(result))
        except Exception as e:
            self.root.after(0, lambda: self.show_error(str(e)))
    
    def show_success(self, input_file, output_file, df, time_columns, converted_counts):
        name_col = 'Name' if 'Name' in df.columns else df.columns[0]
        login_col = 'Login' if 'Login' in df.columns else df.columns[1]
        
        # Полный вывод для копирования
        full = f"{'='*80}\n✅ CONVERSION SUCCESSFUL\n{'='*80}\n"
        full += f"📁 Input: {os.path.basename(input_file)}\n💾 Output: {os.path.basename(output_file)}\n\n"
        full += "⏱️  TIME CONVERSION:\n"
        for col, count in converted_counts.items():
            full += f"   • {col}: {count} entries\n"
        full += f"\n📊 ALL RECORDS ({len(df)} total):\n{'-'*80}\n"
        full += f"{name_col:<25} | {login_col:<15}"
        if 'Entry time' in df.columns:
            full += " | Entry time (UTC+3)"
        if 'Exit time' in df.columns:
            full += " | Exit time (UTC+3)"
        full += f"\n{'-'*80}\n"
        
        for _, row in df.iterrows():
            full += f"{str(row[name_col])[:25]:<25} | {str(row[login_col])[:15]:<15}"
            if 'Entry time' in df.columns:
                full += f" | {row['Entry time']}"
            if 'Exit time' in df.columns:
                full += f" | {row['Exit time']}"
            full += "\n"
        
        full += f"{'-'*80}\n📊 TOTAL: {len(df)} records\n"
        self.full_data = full
        
        # Отображение
        display = full if len(df) <= 50 else full[:5000] + "\n... (full data available via Copy All)\n"
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, display)
        
        self.records_var.set(f"📊 {len(df)} rec")
        self.convert_btn.config(state=tk.NORMAL, bg="#4CD964")
        self.copy_btn.config(state=tk.NORMAL)
        self.copy_vis_btn.config(state=tk.NORMAL)
        self.status_var.set(f"✓ Complete! {len(df)} records")
        
        messagebox.showinfo("Success", f"✅ Conversion completed!\n\n📊 Records: {len(df)}\n💾 Saved: {os.path.basename(output_file)}")
    
    def show_error(self, error_msg):
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, f"❌ ERROR:\n{error_msg}")
        self.convert_btn.config(state=tk.NORMAL, bg="#4CD964")
        self.status_var.set("❌ Conversion failed")
        messagebox.showerror("Error", error_msg)
    
    def copy_all(self):
        if self.full_data:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.full_data)
            self.status_var.set("✅ All data copied!")
            self.copy_btn.config(text="✅ Copied!", bg="#4CAF50")
            self.root.after(2000, lambda: self.copy_btn.config(text="📋 Copy All", bg="#4CAF50"))
    
    def copy_visible(self):
        text = self.result_text.get(1.0, tk.END)
        if text.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_var.set("✅ Visible data copied!")
            self.copy_vis_btn.config(text="✅ Copied!", bg="#4CAF50")
            self.root.after(2000, lambda: self.copy_vis_btn.config(text="📋 Copy Visible", bg="#E8E8E8"))
    
    def clear_results(self):
        self.result_text.delete(1.0, tk.END)
        self.full_data = ""
        self.copy_btn.config(state=tk.DISABLED)
        self.copy_vis_btn.config(state=tk.DISABLED)
        self.records_var.set("")
        self.status_var.set("Ready")

def main():
    root = tk.Tk()
    root.withdraw()
    app = ConverterApp(root)
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
    main()
