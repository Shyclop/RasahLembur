import json
import os
import random
import re
import subprocess
import tempfile
import threading
import time
import urllib.request
from datetime import date, datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "RaSahLembur"
APP_VERSION = "1.1"
GITHUB_REPO = "Shyclop/RasahLembur"
UPDATE_CHECK_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    from openpyxl import load_workbook
    from openpyxl.utils.cell import column_index_from_string, get_column_letter
except ImportError:
    load_workbook = None
    column_index_from_string = None
    get_column_letter = None

try:
    from pynput import keyboard
except ImportError:
    keyboard = None

# -------- BASIC CONFIG --------
WAIT_SECONDS = 3
TYPE_DELAY = 0.02
DEFAULT_LAST_COLUMN = ""
# ------------------------------

class ExcelAutomationUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Automation")
        self.root.geometry("560x800")
        self.root.resizable(True, True)
        
        self.wb = None
        self.sheet = None
        self.start_row = 1
        self.excel_path = tk.StringVar()
        self.current_sheet = tk.StringVar()
        self.input_method = tk.StringVar(value="clipboard")
        self.input_method_display = tk.StringVar(value="Clipboard paste")
        self.randomize_delay = tk.BooleanVar(value=False)
        self.min_delay = tk.DoubleVar(value=0.02)
        self.max_delay = tk.DoubleVar(value=0.2)
        self.end_column = tk.StringVar(value=DEFAULT_LAST_COLUMN)
        self.end_row = tk.IntVar(value=1)
        self.action_steps = ["ctrl+a", "write", "tab"]
        self.action_sequence = tk.StringVar(value="ctrl+a;write;tab")
        self.recording = False
        self.record_listener = None
        self.active_modifiers = set()
        self.keyboard_module = keyboard
        
        # Automation state control
        self.automation_running = False
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self.pause_event.set()  # Not paused initially
        self.resume_countdown = tk.IntVar(value=0)
        
        self.current_language = "en"
        self.language_options = [("English", "en"), ("Bahasa Indonesia", "id")]
        self.translations = {
            "en": {
                "window_title": "Excel Automation",
                "language": "Language:",
                "switch_language": "Switch to Bahasa Indonesia",
                "browse": "Browse",
                "excel_file": "Excel File:",
                "select_sheet": "Select Sheet:",
                "starting_row": "Starting Row:",
                "last_row": "Last Row:",
                "starting_column": "Starting Column:",
                "last_column": "Last Column:",
                "randomize_delays": "Randomize Delays:",
                "input_method_label": "Input Method:",
                "input_method_type": "Type directly",
                "input_method_clipboard": "Clipboard paste",
                "enable_random_delay": "Enable random key press delays",
                "delay_range": "Delay Range (seconds):",
                "min": "Min:",
                "max": "Max:",
                "actions_between_values": "Actions between values:",
                "add": "Add",
                "record": "Record",
                "delete": "Delete",
                "data_preview": "Data Preview:",
                "status_prefix": "Status: {msg}",
                "status_ready": "Status: Ready",
                "execute": "Execute Automation",
                "preview_data": "Preview Data",
                "stop": "Stop",
                "reset": "Reset",
                "info": "Info",
                "record_button": "Record",
                "stop_recording": "Stop Recording",
                "tutorial_title": "Tutorial - User Guide",
                "tutorial_1_title": "📄 Select Excel File",
                "tutorial_1_desc": "Click the Browse button to select an Excel file to process. The selected file will load automatically.",
                "tutorial_2_title": "📋 Select Sheet",
                "tutorial_2_desc": "After the file is loaded, choose the sheet you want to use from the Select Sheet dropdown.",
                "tutorial_3_title": "🔢 Set Rows and Columns",
                "tutorial_3_desc": "Starting Row: the first row to process\nLast Row: the last row to process\nStarting Column: the first column letter (for example A)\nLast Column: the last column letter (for example Z)",
                "tutorial_4_title": "⏱️ Randomize Delays",
                "tutorial_4_desc": "Enable the checkbox to add random delay between steps.\nMin: minimum delay in seconds\nMax: maximum delay in seconds\nThe delay will be applied after each action in the stack.",
                "tutorial_4b_title": "⌨️ Input Method",
                "tutorial_4b_desc": "Choose how each value is inserted into the target app:\n• Type directly - uses pyautogui to type text character by character\n• Clipboard paste - copies the value to the clipboard and sends Ctrl+V\nThis is useful when direct typing fails in some apps or forms.",
                "tutorial_5_title": "🎬 Actions Between Values",
                "tutorial_5_desc": "The list of actions to run for each value from Excel.\nSupported actions:\n• write - type the Excel value\n• tab - press the Tab key\n• enter - press Enter\n• ctrl+a - keyboard shortcut\n• wait:0.5 - wait 0.5 seconds\n• press:delete - press Delete",
                "tutorial_6_title": "➕ Add Button",
                "tutorial_6_desc": "Type an action in the text field and click Add to append it to the stack.",
                "tutorial_7_title": "🎙️ Record Button",
                "tutorial_7_desc": "Click Record to capture a keyboard shortcut or key. Press the desired combination and it will be saved automatically.",
                "tutorial_8_title": "🗑️ Delete Button",
                "tutorial_8_desc": "Select an action from the list and click Delete to remove it.",
                "tutorial_9_title": "👁️ Preview Data",
                "tutorial_9_desc": "Click Preview Data to see a preview of the selected range before automating it.",
                "tutorial_10_title": "▶️ Execute Automation",
                "tutorial_10_desc": "Click to start the automation. The app will count down for 3 seconds, so make sure your target window is active before the timer ends.",
                "tutorial_11_title": "⏹️ Stop",
                "tutorial_11_desc": "Click to stop automation while it is running.",
                "tutorial_12_title": "🔄 Reset",
                "tutorial_12_desc": "Click to reset the app and reload the selected Excel file. This is useful if the file changed in Excel.",
                "warning_title": "Warning",
                "error_title": "Error",
                "info_title": "Info",
                "please_load_sheet": "Please load an Excel file and select a sheet first.",
                "please_enter_starting_column": "Please enter a starting column letter.",
                "please_enter_last_column": "Please enter a last column letter.",
                "invalid_starting_row": "Invalid starting row.",
                "invalid_last_row": "Invalid last row.",
                "invalid_column_range": "Invalid column or sheet data: {error}",
                "invalid_column_or_sheet": "Invalid column range: {error}",
                "no_data_found_range": "No data found from {start_col}{start_row} to {end_col}{end_row}.",
                "preview_summary": "Preview: {count} row(s) from {start_col} to {end_col}",
                "loaded_file": "Loaded: {file_name}",
                "error_loading_file": "Error loading file",
                "no_data_found": "No data found",
                "starting_automation": "Starting automation: {count} row(s) from {start_col} to {end_col}...",
                "click_target_app": "Click target app... ({remaining}s)",
                "automating": "Automating...",
                "resuming": "Resuming... ({remaining}s)",
                "automated_cells": "Automated {count}/{total} cells",
                "automation_complete": "Automation complete!",
                "automation_error": "Error: {error}",
                "stopped": "Stopped",
                "stopping": "Stopping...",
                "paused": "Paused",
                "resuming_in": "Resuming in 3s...",
                "reset_status": "Reset",
                "recording": "Recording... press one key or shortcut",
                "recorded_action": "Recorded action",
                "keyboard_unavailable": "The keyboard recorder is not available in this environment.",
            },
            "id": {
                "window_title": "Otomatisasi Excel",
                "language": "Bahasa:",
                "switch_language": "Ganti ke English",
                "browse": "Jelajahi",
                "excel_file": "Berkas Excel:",
                "select_sheet": "Pilih Lembar:",
                "starting_row": "Baris Awal:",
                "last_row": "Baris Akhir:",
                "last_column": "Kolom Akhir:",
                "starting_column": "Kolom Awal:",
                "randomize_delays": "Acak Jeda:",
                "input_method_label": "Metode Input:",
                "input_method_type": "Ketik langsung",
                "input_method_clipboard": "Tempel clipboard",
                "enable_random_delay": "Aktifkan jeda acak saat menekan tombol",
                "delay_range": "Rentang Jeda (detik):",
                "min": "Min:",
                "max": "Maks:",
                "actions_between_values": "Aksi antar nilai:",
                "add": "Tambah",
                "record": "Rekam",
                "delete": "Hapus",
                "data_preview": "Pratinjau Data:",
                "status_prefix": "Status: {msg}",
                "status_ready": "Status: Siap",
                "execute": "Jalankan Otomatisasi",
                "preview_data": "Pratinjau Data",
                "stop": "Berhenti",
                "reset": "Atur Ulang",
                "info": "Info",
                "record_button": "Rekam",
                "stop_recording": "Berhenti Merekam",
                "tutorial_title": "Tutorial - Panduan Penggunaan",
                "tutorial_1_title": "📄 Pilih File Excel",
                "tutorial_1_desc": "Klik tombol Browse untuk memilih file Excel yang ingin diproses. File yang dipilih akan dimuat otomatis.",
                "tutorial_2_title": "📋 Pilih Lembar",
                "tutorial_2_desc": "Setelah file dimuat, pilih lembar yang ingin digunakan dari dropdown Select Sheet.",
                "tutorial_3_title": "🔢 Atur Baris dan Kolom",
                "tutorial_3_desc": "Baris Awal: baris pertama yang akan diproses\nBaris Akhir: baris terakhir yang akan diproses\nKolom Awal: huruf kolom pertama (misalnya A)\nKolom Akhir: huruf kolom terakhir (misalnya Z)",
                "tutorial_4_title": "⏱️ Acak Jeda",
                "tutorial_4_desc": "Aktifkan checkbox untuk menambahkan jeda acak antar langkah.\nMin: jeda minimum dalam detik\nMax: jeda maksimum dalam detik\nJeda akan diterapkan setelah setiap aksi dalam daftar.",
                "tutorial_4b_title": "⌨️ Metode Input",
                "tutorial_4b_desc": "Pilih cara setiap nilai dimasukkan ke aplikasi target:\n• Ketik langsung - menggunakan pyautogui untuk mengetik teks per karakter\n• Tempel clipboard - menyalin nilai ke clipboard lalu mengirim Ctrl+V\nFitur ini berguna bila metode mengetik langsung tidak bekerja pada beberapa aplikasi atau form.",
                "tutorial_5_title": "🎬 Aksi Antar Nilai",
                "tutorial_5_desc": "Daftar aksi yang akan dijalankan untuk setiap nilai dari Excel.\nAksi yang didukung:\n• write - ketik nilai Excel\n• tab - tekan tombol Tab\n• enter - tekan Enter\n• ctrl+a - kombinasi tombol\n• wait:0.5 - tunggu 0.5 detik\n• press:delete - tekan Delete",
                "tutorial_6_title": "➕ Tombol Tambah",
                "tutorial_6_desc": "Ketik aksi di kolom teks lalu klik Tambah untuk menambahkannya ke daftar.",
                "tutorial_7_title": "🎙️ Tombol Rekam",
                "tutorial_7_desc": "Klik Rekam untuk merekam kombinasi tombol atau tombol tunggal. Tekan kombinasi yang diinginkan dan akan tersimpan otomatis.",
                "tutorial_8_title": "🗑️ Tombol Hapus",
                "tutorial_8_desc": "Pilih aksi dari daftar lalu klik Hapus untuk menghapusnya.",
                "tutorial_9_title": "👁️ Pratinjau Data",
                "tutorial_9_desc": "Klik Pratinjau Data untuk melihat pratinjau rentang data sebelum mengotomatiskan.",
                "tutorial_10_title": "▶️ Jalankan Otomatisasi",
                "tutorial_10_desc": "Klik untuk memulai otomasi. Aplikasi akan menghitung mundur 3 detik, jadi pastikan jendela target sudah aktif sebelum timer habis.",
                "tutorial_11_title": "⏹️ Berhenti",
                "tutorial_11_desc": "Klik untuk menghentikan otomasi saat sedang berjalan.",
                "tutorial_12_title": "🔄 Atur Ulang",
                "tutorial_12_desc": "Klik untuk mengatur ulang aplikasi dan memuat ulang file Excel yang dipilih. Berguna jika file telah berubah di Excel.",
                "warning_title": "Peringatan",
                "error_title": "Kesalahan",
                "info_title": "Info",
                "please_load_sheet": "Silakan muat file Excel dan pilih lembar terlebih dahulu.",
                "please_enter_starting_column": "Silakan masukkan huruf kolom awal.",
                "please_enter_last_column": "Silakan masukkan huruf kolom akhir.",
                "invalid_starting_row": "Baris awal tidak valid.",
                "invalid_last_row": "Baris akhir tidak valid.",
                "invalid_column_range": "Kolom atau data lembar tidak valid: {error}",
                "invalid_column_or_sheet": "Rentang kolom tidak valid: {error}",
                "no_data_found_range": "Tidak ada data dari {start_col}{start_row} hingga {end_col}{end_row}.",
                "preview_summary": "Pratinjau: {count} baris dari {start_col} hingga {end_col}",
                "loaded_file": "Dimuat: {file_name}",
                "error_loading_file": "Gagal memuat file",
                "no_data_found": "Tidak ada data",
                "starting_automation": "Memulai otomatisasi: {count} baris dari {start_col} hingga {end_col}...",
                "click_target_app": "Klik aplikasi target... ({remaining}s)",
                "automating": "Mengotomatisasi...",
                "resuming": "Melanjutkan... ({remaining}s)",
                "automated_cells": "Terotomatisasi {count}/{total} sel",
                "automation_complete": "Otomatisasi selesai!",
                "automation_error": "Kesalahan: {error}",
                "stopped": "Dihentikan",
                "stopping": "Menghentikan...",
                "paused": "Dijeda",
                "resuming_in": "Melanjutkan dalam 3 detik...",
                "reset_status": "Reset",
                "recording": "Merekam... tekan satu tombol atau shortcut",
                "recorded_action": "Aksi terekam",
                "keyboard_unavailable": "Perekam keyboard tidak tersedia di lingkungan ini.",
            }
        }
        self.input_method_display.set(self._translate("input_method_clipboard"))
        self.translatable_widgets = {}
        
        self.create_widgets()
        self.apply_language()
        
    def register_text_widget(self, widget, key):
        self.translatable_widgets[key] = widget

    def _translate(self, key, **kwargs):
        translation = self.translations[self.current_language].get(key)
        if translation is None:
            translation = self.translations["en"].get(key, key)
        if kwargs:
            return translation.format(**kwargs)
        return translation

    def on_language_selected(self, event=None):
        selected_name = self.language_var.get()
        for display_name, code in self.language_options:
            if display_name == selected_name:
                self.current_language = code
                self.apply_language()
                break

    def on_input_method_selected(self, event=None):
        display_value = self.input_method_display.get()
        allowed = {
            self._translate("input_method_type"): "type",
            self._translate("input_method_clipboard"): "clipboard",
        }
        self.input_method.set(allowed.get(display_value, "clipboard"))
        self.update_status(self._translate("status_ready"), "blue")

    def apply_language(self):
        self.root.title(self._translate("window_title"))
        self.language_label.config(text=self._translate("language"))

        current_display = next(
            (display_name for display_name, code in self.language_options if code == self.current_language),
            "English",
        )
        self.language_var.set(current_display)

        for key, widget in self.translatable_widgets.items():
            if isinstance(widget, (ttk.Label, ttk.Button, ttk.Checkbutton)):
                widget.config(text=self._translate(key))

        if hasattr(self, "input_method_dropdown"):
            self.input_method_dropdown['values'] = [
                self._translate("input_method_type"),
                self._translate("input_method_clipboard"),
            ]
            current_method = self.input_method.get()
            display_value = self._translate("input_method_type") if current_method == "type" else self._translate("input_method_clipboard")
            self.input_method_display.set(display_value)
            self.input_method_dropdown.set(display_value)

        if hasattr(self, "status_label"):
            self.status_label.config(text=self._translate("status_ready"))

    def _require_excel_support(self):
        if load_workbook is None or column_index_from_string is None or get_column_letter is None:
            messagebox.showerror(
                self._translate("error_title"),
                "Excel support is unavailable because the required local dependency is missing."
            )
            self.update_status("Excel support unavailable", "red")
            return False
        return True

    def _require_automation_support(self):
        if pyautogui is None:
            messagebox.showerror(
                self._translate("error_title"),
                "Automation is unavailable because the required local dependency is missing."
            )
            self.update_status("Automation support unavailable", "red")
            return False
        return True

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)

        # Top bar
        top_bar = ttk.Frame(self.root, padding="10 10 0 10")
        top_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.language_label = ttk.Label(top_bar, text="Language:")
        self.language_label.grid(row=0, column=0, sticky=tk.W)
        self.language_var = tk.StringVar(value=self.current_language)
        self.language_dropdown = ttk.Combobox(
            top_bar,
            textvariable=self.language_var,
            state="readonly",
            width=20,
            values=[name for name, _ in self.language_options],
        )
        self.language_dropdown.grid(row=0, column=1, sticky=tk.E)
        self.language_dropdown.bind("<<ComboboxSelected>>", self.on_language_selected)

        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=0)
        
        # ===== FILE SELECTION =====
        file_label = ttk.Label(main_frame, text="Excel File:", font=("Arial", 11, "bold"))
        self.register_text_widget(file_label, "excel_file")
        file_label.grid(row=1, column=0, columnspan=1, sticky=tk.W, pady=(0, 5))
        
        file_entry = ttk.Entry(main_frame, textvariable=self.excel_path, width=50)
        file_entry.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_file)
        self.register_text_widget(browse_btn, "browse")
        browse_btn.grid(row=2, column=1, sticky=tk.E)
        
        # ===== SHEET SELECTION =====
        sheet_label = ttk.Label(main_frame, text="Select Sheet:", font=("Arial", 11, "bold"))
        self.register_text_widget(sheet_label, "select_sheet")
        sheet_label.grid(row=3, column=0, columnspan=2, sticky=tk.W)
        
        self.sheet_dropdown = ttk.Combobox(main_frame, textvariable=self.current_sheet, state="readonly", width=58)
        self.sheet_dropdown.grid(row=4, column=0, columnspan=1, sticky=(tk.W, tk.E))
        self.sheet_dropdown.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_selected())
        
        # ===== START/END ROW =====
        row_label = ttk.Label(main_frame, text="Starting Row:", font=("Arial", 11, "bold"))
        self.register_text_widget(row_label, "starting_row")
        row_label.grid(row=5, column=0, columnspan=2, sticky=tk.W)
        
        self.row_spinbox = ttk.Spinbox(main_frame, from_=1, to=10000, width=12)
        self.row_spinbox.set(1)
        self.row_spinbox.grid(row=6, column=0, columnspan=2, sticky=tk.W)
        
        end_row_label = ttk.Label(main_frame, text="Last Row:", font=("Arial", 11, "bold"))
        self.register_text_widget(end_row_label, "last_row")
        end_row_label.grid(row=7, column=0, columnspan=2, sticky=tk.W)
        
        self.end_row_spinbox = ttk.Spinbox(main_frame, from_=1, to=10000, width=12, textvariable=self.end_row)
        self.end_row_spinbox.set(1)
        self.end_row_spinbox.grid(row=8, column=0, columnspan=2, sticky=tk.W)
        
        # ===== START/END COLUMN =====
        col_label = ttk.Label(main_frame, text="Starting Column:", font=("Arial", 11, "bold"))
        self.register_text_widget(col_label, "starting_column")
        col_label.grid(row=9, column=0, columnspan=2, sticky=tk.W)
        
        self.col_entry = ttk.Entry(main_frame, width=12)
        self.col_entry.grid(row=10, column=0, columnspan=2, sticky=tk.W)
        
        end_col_label = ttk.Label(main_frame, text="Last Column:", font=("Arial", 11, "bold"))
        self.register_text_widget(end_col_label, "last_column")
        end_col_label.grid(row=11, column=0, columnspan=2, sticky=tk.W)
        
        self.end_col_entry = ttk.Entry(main_frame, width=12, textvariable=self.end_column)
        self.end_col_entry.grid(row=12, column=0, columnspan=2, sticky=tk.W)
        
        # ===== INPUT METHOD =====
        input_method_label = ttk.Label(main_frame, text="Input Method:", font=("Arial", 11, "bold"))
        self.register_text_widget(input_method_label, "input_method_label")
        input_method_label.grid(row=13, column=0, columnspan=2, sticky=tk.W)

        self.input_method_dropdown = ttk.Combobox(
            main_frame,
            textvariable=self.input_method_display,
            state="readonly",
            width=22,
            values=[self._translate("input_method_type"), self._translate("input_method_clipboard")],
        )
        self.input_method_dropdown.grid(row=14, column=0, columnspan=2, sticky=tk.W)
        self.input_method_dropdown.bind("<<ComboboxSelected>>", self.on_input_method_selected)

        # ===== RANDOMIZE DELAY =====
        randomize_label = ttk.Label(main_frame, text="Randomize Delays:", font=("Arial", 11, "bold"))
        self.register_text_widget(randomize_label, "randomize_delays")
        randomize_label.grid(row=15, column=0, columnspan=1, sticky=tk.W)
        
        self.randomize_check = ttk.Checkbutton(main_frame, text="Enable random key press delays", variable=self.randomize_delay)
        self.register_text_widget(self.randomize_check, "enable_random_delay")
        self.randomize_check.grid(row=16, column=0, sticky=tk.W)
        
        delay_range_label = ttk.Label(main_frame, text="Delay Range (seconds):", font=("Arial", 10))
        self.register_text_widget(delay_range_label, "delay_range")
        delay_range_label.grid(row=17, column=0, sticky=tk.W, pady=(0, 5))
        
        delay_frame = ttk.Frame(main_frame)
        delay_frame.grid(row=18, column=0, sticky=tk.W, pady=(0, 10))
        
        min_label = ttk.Label(delay_frame, text="Min:")
        self.register_text_widget(min_label, "min")
        min_label.pack(side=tk.LEFT, padx=(0, 5))
        self.min_delay_spinbox = ttk.Spinbox(delay_frame, from_=0.01, to=5.0, width=8, textvariable=self.min_delay, format="%.2f")
        self.min_delay_spinbox.pack(side=tk.LEFT, padx=(0, 15))
        
        max_label = ttk.Label(delay_frame, text="Max:")
        self.register_text_widget(max_label, "max")
        max_label.pack(side=tk.LEFT, padx=(0, 5))
        self.max_delay_spinbox = ttk.Spinbox(delay_frame, from_=0.01, to=5.0, width=8, textvariable=self.max_delay, format="%.2f")
        self.max_delay_spinbox.pack(side=tk.LEFT)

        # ===== ACTION STACK =====
        action_label = ttk.Label(main_frame, text="Actions between values:", font=("Arial", 11, "bold"))
        self.register_text_widget(action_label, "actions_between_values")
        action_label.grid(row=19, column=0, sticky=tk.W, pady=(10, 5))

        action_controls = ttk.Frame(main_frame)
        action_controls.grid(row=20, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

        self.action_entry = ttk.Entry(action_controls, textvariable=self.action_sequence, width=48)
        self.action_entry.pack(side=tk.LEFT, padx=(0, 5))

        self.add_action_btn = ttk.Button(action_controls, text="Add", command=self.add_action_step)
        self.register_text_widget(self.add_action_btn, "add")
        self.add_action_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.record_btn = ttk.Button(action_controls, text="Record", command=self.toggle_recording)
        self.register_text_widget(self.record_btn, "record")
        self.record_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.delete_action_btn = ttk.Button(action_controls, text="Delete", command=self.delete_selected_action)
        self.register_text_widget(self.delete_action_btn, "delete")
        self.delete_action_btn.pack(side=tk.LEFT)

        action_list_frame = ttk.Frame(main_frame)
        action_list_frame.grid(row=21, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        action_scrollbar = ttk.Scrollbar(action_list_frame)
        action_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.action_steps_listbox = tk.Listbox(action_list_frame, height=5, yscrollcommand=action_scrollbar.set, selectmode=tk.SINGLE)
        self.action_steps_listbox.pack(side=tk.LEFT, fill=(tk.BOTH), expand=True)
        action_scrollbar.config(command=self.action_steps_listbox.yview)
        self.refresh_action_list()
        
        # ===== DATA PREVIEW =====
        preview_label = ttk.Label(main_frame, text="Data Preview:", font=("Arial", 11, "bold"))
        self.register_text_widget(preview_label, "data_preview")
        preview_label.grid(row=22, column=0, sticky=tk.W, pady=(15, 5))
        
        # Listbox with scrollbar
        frame_listbox = ttk.Frame(main_frame)
        frame_listbox.grid(row=23, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(frame_listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_listbox = tk.Listbox(frame_listbox, height=8, yscrollcommand=scrollbar.set)
        self.preview_listbox.pack(side=tk.LEFT, fill=(tk.BOTH), expand=True)
        scrollbar.config(command=self.preview_listbox.yview)
        
        # ===== STATUS =====
        self.status_label = ttk.Label(main_frame, text="Status: Ready", foreground="blue")
        self.status_label.grid(row=24, column=0, sticky=tk.W, pady=(10, 0))
        
        # ===== BUTTONS =====
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=25, column=0, sticky=(tk.W, tk.E), pady=(15, 0))
        
        self.execute_btn = ttk.Button(button_frame, text="Execute Automation", command=self.execute_automation)
        self.register_text_widget(self.execute_btn, "execute")
        self.execute_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.update_btn = ttk.Button(button_frame, text="Check Update", command=self.check_for_updates)
        self.update_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.preview_btn = ttk.Button(button_frame, text="Preview Data", command=self.preview_data)
        self.register_text_widget(self.preview_btn, "preview_data")
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_automation, state=tk.DISABLED)
        self.register_text_widget(self.stop_btn, "stop")
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.reset_btn = ttk.Button(button_frame, text="Reset", command=self.reset_automation)
        self.register_text_widget(self.reset_btn, "reset")
        self.reset_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.info_btn = ttk.Button(button_frame, text="Info", command=self.show_tutorial)
        self.register_text_widget(self.info_btn, "info")
        self.info_btn.pack(side=tk.LEFT)

        self.apply_language()
        
    def browse_file(self):
        if not self._require_excel_support():
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if file_path:
            self.excel_path.set(file_path)
            self.load_workbook(file_path)
            
    def load_workbook(self, path):
        if not self._require_excel_support():
            return

        try:
            self.wb = load_workbook(path, data_only=True)
            self.sheet_dropdown['values'] = self.wb.sheetnames
            current_sheet_name = self.current_sheet.get()
            if current_sheet_name in self.wb.sheetnames:
                self.sheet_dropdown.set(current_sheet_name)
            elif self.wb.sheetnames:
                self.sheet_dropdown.current(0)
            self.on_sheet_selected()
            self.update_status(self._translate("loaded_file", file_name=os.path.basename(path)), "green")
        except Exception as e:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_column_range", error=str(e)))
            self.update_status(self._translate("error_loading_file"), "red")
            
    def on_sheet_selected(self):
        if self.wb:
            sheet_name = self.current_sheet.get()
            self.sheet = self.wb[sheet_name]
            self.preview_listbox.delete(0, tk.END)
            
    def get_columns(self, start_column, end_column=None):
        if not self._require_excel_support():
            raise RuntimeError("Excel support unavailable")

        if end_column is None:
            end_column = self.end_column.get()
        start_idx = column_index_from_string(start_column)
        end_idx = column_index_from_string(end_column)
        if end_idx < start_idx:
            raise ValueError("Last column must be the same or to the right of the starting column.")
        return [get_column_letter(i) for i in range(start_idx, end_idx + 1)]

    def format_cell_value(self, cell):
        value = cell.value
        if value is None:
            return ""

        if isinstance(value, datetime):
            number_format = getattr(cell, "number_format", None) or ""
            if number_format and any(token in number_format.lower() for token in ["h", "m", "s"]):
                return value.date().strftime("%Y-%m-%d")

            try:
                return self._format_excel_date(value.date(), number_format)
            except Exception:
                return value.date().strftime("%Y-%m-%d")

        if isinstance(value, date):
            number_format = getattr(cell, "number_format", None) or ""
            try:
                return self._format_excel_date(value, number_format)
            except Exception:
                return value.strftime("%Y-%m-%d")

        return str(value)

    def _format_excel_date(self, value, number_format):
        if not number_format:
            return value.strftime("%Y-%m-%d")

        fmt = number_format.strip().split(";")[0]
        fmt = fmt.replace("[$-409]", "").replace("[$-F800]", "")
        fmt = fmt.replace(" ", "")

        if any(token in fmt.lower() for token in ["h", "s", "am/pm", "a/p"]):
            return value.strftime("%Y-%m-%d")

        replacements = [
            ("yyyy", "%Y"),
            ("yy", "%y"),
            ("mmmm", "%B"),
            ("mmm", "%b"),
            ("mm", "%m"),
            ("m", "%m"),
            ("dddd", "%A"),
            ("ddd", "%a"),
            ("dd", "%d"),
            ("d", "%d"),
        ]

        formatted = fmt
        for excel_token, python_token in replacements:
            formatted = formatted.replace(excel_token, python_token)

        if formatted == fmt:
            return value.strftime("%Y-%m-%d")

        return value.strftime(formatted)

    def read_rows(self, start_column, start_row, end_column=None, end_row=None):
        if end_row is None:
            end_row = self.end_row.get()
        if end_row < start_row:
            raise ValueError("Last row must be the same or greater than the starting row.")
        columns = self.get_columns(start_column, end_column)
        rows = []
        for row in range(start_row, end_row + 1):
            values = []
            for col in columns:
                cell = self.sheet[f"{col}{row}"]
                values.append(self.format_cell_value(cell))
            rows.append(values)
        return rows

    def preview_data(self):
        if not self._require_excel_support():
            return

        if not self.sheet:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_load_sheet"))
            return
            
        col = self.col_entry.get().strip().upper()
        if not col:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_enter_starting_column"))
            return
            
        end_col = self.end_column.get().strip().upper()
        if not end_col:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_enter_last_column"))
            return
            
        try:
            start_row = int(self.row_spinbox.get())
        except ValueError:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_starting_row"))
            return
            
        try:
            end_row = int(self.end_row_spinbox.get())
        except ValueError:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_last_row"))
            return
            
        try:
            rows = self.read_rows(col, start_row, end_col, end_row)
        except Exception as e:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_column_range", error=str(e)))
            return
            
        self.preview_listbox.delete(0, tk.END)
        for row_values in rows:
            self.preview_listbox.insert(tk.END, " | ".join(row_values))
            
        if not rows:
            messagebox.showinfo(self._translate("info_title"), self._translate("no_data_found_range", start_col=col, start_row=start_row, end_col=end_col, end_row=end_row))
            self.update_status(self._translate("no_data_found"), "orange")
        else:
            self.update_status(self._translate("preview_summary", count=len(rows), start_col=col, end_col=end_col), "blue")
            
    def execute_automation(self):
        if not self._require_automation_support():
            return

        if not self.sheet:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_load_sheet"))
            return
            
        col = self.col_entry.get().strip().upper()
        if not col:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_enter_starting_column"))
            return
            
        end_col = self.end_column.get().strip().upper()
        if not end_col:
            messagebox.showwarning(self._translate("warning_title"), self._translate("please_enter_last_column"))
            return
            
        try:
            start_row = int(self.row_spinbox.get())
        except ValueError:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_starting_row"))
            return
            
        try:
            end_row = int(self.end_row_spinbox.get())
        except ValueError:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_last_row"))
            return
            
        if end_row < start_row:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_last_row"))
            return
            
        try:
            self.get_columns(col, end_col)
        except Exception as e:
            messagebox.showerror(self._translate("error_title"), self._translate("invalid_column_or_sheet", error=str(e)))
            return
            
        # Disable button during execution
        self.execute_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.automation_running = True
        self.pause_event.set()
        self.stop_event.clear()
        
        # Run in separate thread to prevent UI freezing
        thread = threading.Thread(target=self.run_automation, args=(col, start_row, end_row))
        thread.daemon = True
        thread.start()
        
    def pause_automation(self):
        if self.automation_running:
            if self.pause_event.is_set():
                # Pause
                self.pause_event.clear()
                self.pause_btn.config(text="Resume")
                self.update_status(self._translate("paused"), "orange")
            else:
                # Resume
                self.pause_event.set()
                self.pause_btn.config(text="Pause")
                self.resume_countdown.set(3)  # 3 second countdown
                self.update_status(self._translate("resuming_in"), "blue")
    
    def stop_automation(self):
        self.stop_event.set()
        self.pause_event.set()  # Ensure not paused so it can exit
        self.update_status(self._translate("stopping"), "orange")
    
    def reset_automation(self):
        self.stop_event.set()
        self.pause_event.set()
        self.resume_countdown.set(0)
        self.automation_running = False
        self.execute_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.preview_listbox.delete(0, tk.END)

        current_file = self.excel_path.get().strip()
        if current_file and os.path.exists(current_file):
            self.load_workbook(current_file)
        else:
            self.update_status(self._translate("reset_status"), "blue")

    def refresh_action_list(self):
        self.action_steps_listbox.delete(0, tk.END)
        for action in self.action_steps:
            self.action_steps_listbox.insert(tk.END, action)
        if self.action_steps:
            selected_index = self.action_steps_listbox.curselection()
            if not selected_index:
                self.action_steps_listbox.selection_set(0)
                self.action_steps_listbox.activate(0)
            else:
                self.action_steps_listbox.selection_set(selected_index[0])
                self.action_steps_listbox.activate(selected_index[0])

    def add_action_step(self, action=None):
        if action is None:
            action = self.action_sequence.get().strip()

        if action:
            self.action_steps.append(action)
            self.refresh_action_list()
            self.action_sequence.set("")

    def delete_selected_action(self):
        selected_index = self.action_steps_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        if 0 <= index < len(self.action_steps):
            del self.action_steps[index]
            self.refresh_action_list()
            if self.action_steps:
                target_index = min(index, len(self.action_steps) - 1)
                self.action_steps_listbox.selection_set(target_index)
                self.action_steps_listbox.activate(target_index)

    def toggle_recording(self):
        if self.keyboard_module is None:
            messagebox.showwarning(self._translate("warning_title"), self._translate("keyboard_unavailable"))
            return

        if self.recording:
            self._stop_recording()
            return

        self.recording = True
        self.active_modifiers.clear()
        self.record_listener = self.keyboard_module.Listener(on_press=self._on_record_key_press, on_release=self._on_record_key_release)
        self.record_listener.start()
        self.record_btn.config(text=self._translate("stop_recording"))
        self.update_status(self._translate("recording"), "orange")

    def _stop_recording(self):
        if self.record_listener is not None:
            self.record_listener.stop()
            self.record_listener = None
        self.recording = False
        self.active_modifiers.clear()
        self.record_btn.config(text=self._translate("record_button"))

    def _on_record_key_press(self, key):
        if not self.recording:
            return

        # If it's a modifier, just track it and wait for another key
        if isinstance(key, self.keyboard_module.Key):
            modifier_name = self._modifier_name(key)
            if modifier_name:
                self.active_modifiers.add(modifier_name)
                return

        # If we have modifiers but no printable char, try to extract key name
        if self.active_modifiers and not (hasattr(key, "char") and key.char is not None and len(key.char) == 1 and ord(key.char) >= 32):
            # For Ctrl+letter combinations, try to get the key name
            try:
                key_name = key.name if hasattr(key, "name") else None
                if not key_name and hasattr(key, "vk"):
                    # Try to convert VK code to letter (65=A, 66=B, etc.)
                    if 65 <= key.vk <= 90:
                        key_name = chr(key.vk).lower()
                    elif 48 <= key.vk <= 57:
                        key_name = chr(key.vk)
                
                if key_name:
                    action = f"{'+'.join(sorted(list(self.active_modifiers)) + [key_name])}"
                    self.root.after(0, self.add_action_step, action)
                    self.root.after(0, self._stop_recording)
                    self.root.after(0, self.update_status, self._translate("recorded_action"), "blue")
                    self.root.after(0, self.active_modifiers.clear)
                    return
            except (AttributeError, TypeError):
                pass

        # Normal key press handling
        action = self._build_recorded_action(key)
        if action:
            self.root.after(0, self.add_action_step, action)
            self.root.after(0, self._stop_recording)
            self.root.after(0, self.update_status, self._translate("recorded_action"), "blue")
            self.root.after(0, self.active_modifiers.clear)

    def _on_record_key_release(self, key):
        if not self.recording:
            return
        if isinstance(key, self.keyboard_module.Key):
            if key in self._modifier_keys():
                self.active_modifiers.discard(self._modifier_name(key))

    def _build_recorded_action(self, key):
        if self.keyboard_module is None:
            return None

        modifiers = [m for m in self.active_modifiers if m]
        key_name = None

        # Try to get a readable key name
        if hasattr(key, "char") and key.char is not None:
            # Handle special control characters
            if key.char == "\n":
                key_name = "enter"
            elif key.char == "\t":
                key_name = "tab"
            elif key.char == "\b":
                key_name = "backspace"
            elif key.char == " ":
                key_name = "space"
            elif len(key.char) == 1 and ord(key.char) >= 32:
                # Regular printable character
                key_name = key.char.lower() if key.char.isalpha() else key.char
            # else: control character, try other methods below
        
        # If no char, try special keys lookup
        if not key_name:
            key_name = self._special_key_name(key)
        
        # Last resort: use the key's name attribute
        if not key_name:
            try:
                key_name = getattr(key, "name", None)
            except (AttributeError, TypeError):
                pass

        if not key_name:
            return None

        if modifiers:
            return f"{'+'.join(sorted(modifiers) + [key_name])}"
        return key_name

    def _modifier_keys(self):
        if self.keyboard_module is None:
            return []
        return [
            self.keyboard_module.Key.ctrl,
            self.keyboard_module.Key.ctrl_l,
            self.keyboard_module.Key.ctrl_r,
            self.keyboard_module.Key.shift,
            self.keyboard_module.Key.shift_l,
            self.keyboard_module.Key.shift_r,
            self.keyboard_module.Key.alt,
            self.keyboard_module.Key.alt_l,
            self.keyboard_module.Key.alt_r,
            self.keyboard_module.Key.cmd,
            self.keyboard_module.Key.cmd_l,
            self.keyboard_module.Key.cmd_r,
        ]

    def _modifier_name(self, key):
        names = {
            self.keyboard_module.Key.ctrl: "ctrl",
            self.keyboard_module.Key.ctrl_l: "ctrl",
            self.keyboard_module.Key.ctrl_r: "ctrl",
            self.keyboard_module.Key.shift: "shift",
            self.keyboard_module.Key.shift_l: "shift",
            self.keyboard_module.Key.shift_r: "shift",
            self.keyboard_module.Key.alt: "alt",
            self.keyboard_module.Key.alt_l: "alt",
            self.keyboard_module.Key.alt_r: "alt",
            self.keyboard_module.Key.cmd: "cmd",
            self.keyboard_module.Key.cmd_l: "cmd",
            self.keyboard_module.Key.cmd_r: "cmd",
        }
        return names.get(key)

    def _special_key_name(self, key):
        names = {
            self.keyboard_module.Key.enter: "enter",
            self.keyboard_module.Key.backspace: "backspace",
            self.keyboard_module.Key.tab: "tab",
            self.keyboard_module.Key.esc: "esc",
            self.keyboard_module.Key.space: "space",
            self.keyboard_module.Key.left: "left",
            self.keyboard_module.Key.right: "right",
            self.keyboard_module.Key.up: "up",
            self.keyboard_module.Key.down: "down",
            self.keyboard_module.Key.home: "home",
            self.keyboard_module.Key.end: "end",
            self.keyboard_module.Key.delete: "delete",
        }
        return names.get(key)

    def write_text(self, text):
        value = "" if text is None else str(text)
        if value == "":
            return

        if self.input_method.get() == "type":
            if pyautogui is None:
                raise RuntimeError("Automation support unavailable")
            pyautogui.write(value, interval=TYPE_DELAY)
            return

        self._paste_text(value)

    def _paste_text(self, text):
        value = "" if text is None else str(text)
        if value == "":
            return

        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(value)
            self.root.update()
        except Exception:
            try:
                import pyperclip
                pyperclip.copy(value)
            except Exception:
                if pyautogui is not None:
                    pyautogui.write(value, interval=TYPE_DELAY)
                return

        time.sleep(0.05)
        if pyautogui is not None:
            pyautogui.hotkey("ctrl", "v")

    def execute_action_sequence(self, text):
        if not self._require_automation_support():
            return False

        actions = [step.strip() for step in self.action_steps if str(step).strip()]
        if not actions:
            actions = ["write"]

        for action in actions:
            if not action:
                continue

            normalized = action.replace(" ", "").lower()
            if normalized in {"write", "type", "value", "{{value}}", "{{data}}"}:
                self.write_text(text)
            elif normalized.startswith("wait:"):
                try:
                    time.sleep(float(normalized.split(":", 1)[1]))
                except ValueError:
                    raise ValueError(f"Invalid wait value: {action}")
            elif normalized.startswith("press:"):
                pyautogui.press(normalized.split(":", 1)[1])
            elif normalized.startswith("hotkey:"):
                keys = [key for key in normalized.split(":", 1)[1].split("+") if key]
                pyautogui.hotkey(*keys)
            elif "+" in normalized:
                keys = [key for key in normalized.split("+") if key]
                pyautogui.hotkey(*keys)
            elif normalized in {"tab", "enter", "backspace", "delete", "left", "right", "up", "down", "esc", "home", "end", "space"}:
                pyautogui.press(normalized)
            elif normalized.startswith("write:"):
                self.write_text(normalized.split(":", 1)[1])
            else:
                self.write_text(action)

            if self.randomize_delay.get():
                time.sleep(random.uniform(self.min_delay.get(), self.max_delay.get()))
            else:
                time.sleep(0.2)

        return True

    def check_for_updates(self):
        self.update_btn.config(state=tk.DISABLED)
        try:
            latest_release = get_latest_release(GITHUB_REPO)
            if not latest_release:
                self.update_status("Update check failed", "orange")
                return

            latest_tag = latest_release.get("tag_name", "")
            asset = find_release_asset(latest_release, APP_NAME)
            if not asset:
                self.update_status(f"No installer found for {latest_tag}", "orange")
                return

            if _version_to_tuple(latest_tag) <= _version_to_tuple(APP_VERSION):
                self.update_status(f"You are on the latest version ({APP_VERSION})", "green")
                return

            download_url = asset.get("browser_download_url")
            if not download_url:
                self.update_status("Update asset URL missing", "orange")
                return

            confirm = messagebox.askyesno(
                "Update available",
                f"A newer version ({latest_tag}) is available.\n\nDo you want to download and install it now?",
            )
            if confirm:
                self.download_and_install_update(download_url, asset.get("name", f"{APP_NAME}Installer.exe"))
        except Exception as exc:
            self.update_status(f"Update check error: {exc}", "red")
        finally:
            self.update_btn.config(state=tk.NORMAL)

    def download_and_install_update(self, download_url, file_name):
        if GITHUB_REPO == "YOUR_USERNAME/YOUR_REPO":
            messagebox.showinfo("Update not configured", "Please set the GitHub repo in the code before enabling automatic updates.")
            return

        temp_dir = tempfile.gettempdir()
        local_path = os.path.join(temp_dir, file_name)
        self.update_status("Downloading update...", "orange")
        try:
            urllib.request.urlretrieve(download_url, local_path)
            messagebox.showinfo("Update ready", f"Installer downloaded to:\n{local_path}\n\nRun it to complete the update.")
            self.update_status("Update downloaded", "green")
            if os.path.exists(local_path):
                subprocess.Popen([local_path], shell=True)
                self.root.after(500, self.root.destroy)
        except Exception as exc:
            self.update_status(f"Download failed: {exc}", "red")
            messagebox.showerror("Update failed", str(exc))

    def run_automation(self, col, start_row, end_row):
        try:
            # Reset countdown at start
            self.resume_countdown.set(0)
            
            # Read row-by-row data from start to end column and row
            end_col = self.end_column.get().strip().upper()
            rows = self.read_rows(col, start_row, end_col, end_row)
            if not rows:
                self.update_status(self._translate("no_data_found"), "orange")
                self.automation_running = False
                self.execute_btn.config(state=tk.NORMAL)
                pause_btn = getattr(self, "pause_btn", None)
                if pause_btn is not None:
                    pause_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
            self.update_status(self._translate("starting_automation", count=len(rows), start_col=col, end_col=end_col), "blue")
            
            # Wait for user to focus target app
            for remaining in range(WAIT_SECONDS, 0, -1):
                if self.stop_event.is_set():
                    self.update_status("Stopped", "red")
                    return
                self.update_status(self._translate("click_target_app", remaining=remaining), "orange")
                time.sleep(1)
                
            # Type into target app row by row
            self.update_status(self._translate("automating"), "blue")
            total_cells = len(rows) * len(self.get_columns(col, end_col))
            cell_count = 0
            for row_values in rows:
                if self.stop_event.is_set():
                    self.update_status("Stopped", "red")
                    return
                    
                for col_index, text in enumerate(row_values):
                    # Check for stop signal
                    if self.stop_event.is_set():
                        self.update_status("Stopped", "red")
                        return
                    
                    # Handle pause
                    while not self.pause_event.is_set():
                        if self.stop_event.is_set():
                            self.update_status("Stopped", "red")
                            return
                        time.sleep(0.1)
                    
                    # Countdown after resume
                    countdown_value = self.resume_countdown.get()
                    if countdown_value > 0:
                        for remaining in range(countdown_value, 0, -1):
                            if self.stop_event.is_set():
                                self.update_status("Stopped", "red")
                                return
                            self.update_status(self._translate("resuming", remaining=remaining), "orange")
                            time.sleep(1)
                        self.resume_countdown.set(0)
                    
                    self.execute_action_sequence(text)
                    cell_count += 1
                    
                    # Delay after each cell if enabled
                    if self.randomize_delay.get():
                        delay = random.uniform(self.min_delay.get(), self.max_delay.get())
                    else:
                        delay = 0.2
                    time.sleep(delay)
                    self.update_status(self._translate("automated_cells", count=cell_count, total=total_cells), "blue")
                
            self.update_status(self._translate("automation_complete"), "green")
        except Exception as e:
            self.update_status(self._translate("automation_error", error=str(e)), "red")
        finally:
            self.automation_running = False
            self.resume_countdown.set(0)
            self.execute_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            
    def update_status(self, msg, color="blue"):
        self.status_label.config(text=self._translate("status_prefix", msg=msg), foreground=color)
        self.root.update()
    
    def show_tutorial(self):
        """Tampilkan tutorial dalam Bahasa Indonesia"""
        tutorial_window = tk.Toplevel(self.root)
        tutorial_window.title(self._translate("tutorial_title"))
        tutorial_window.geometry("650x700")
        
        # Frame utama dengan scrollbar
        canvas_frame = ttk.Frame(tutorial_window)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg="white")
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Konten tutorial
        tutorials = [
            (self._translate("tutorial_1_title"), self._translate("tutorial_1_desc")),
            (self._translate("tutorial_2_title"), self._translate("tutorial_2_desc")),
            (self._translate("tutorial_3_title"), self._translate("tutorial_3_desc")),
            (self._translate("tutorial_4_title"), self._translate("tutorial_4_desc")),
            (self._translate("tutorial_4b_title"), self._translate("tutorial_4b_desc")),
            (self._translate("tutorial_5_title"), self._translate("tutorial_5_desc")),
            (self._translate("tutorial_6_title"), self._translate("tutorial_6_desc")),
            (self._translate("tutorial_7_title"), self._translate("tutorial_7_desc")),
            (self._translate("tutorial_8_title"), self._translate("tutorial_8_desc")),
            (self._translate("tutorial_9_title"), self._translate("tutorial_9_desc")),
            (self._translate("tutorial_10_title"), self._translate("tutorial_10_desc")),
            (self._translate("tutorial_11_title"), self._translate("tutorial_11_desc")),
            (self._translate("tutorial_12_title"), self._translate("tutorial_12_desc")),
        ]
        
        for title, desc in tutorials:
            # Title
            title_label = ttk.Label(
                scrollable_frame, 
                text=title, 
                font=("Arial", 10, "bold"),
                foreground="#0066cc",
                background="white"
            )
            title_label.pack(anchor=tk.W, padx=15, pady=(10, 3))
            
            # Description
            desc_label = ttk.Label(
                scrollable_frame, 
                text=desc, 
                font=("Arial", 9),
                wraplength=600,
                justify=tk.LEFT,
                background="white"
            )
            desc_label.pack(anchor=tk.W, padx=30, pady=(0, 8))
            
            # Separator
            separator = ttk.Separator(scrollable_frame, orient=tk.HORIZONTAL)
            separator.pack(fill=tk.X, padx=15, pady=5)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

def _version_to_tuple(value):
    version_text = str(value or "")
    match = re.search(r"(\d+(?:\.\d+)+)", version_text)
    if not match:
        return (0,)
    return tuple(int(part) for part in match.group(1).split("."))


def find_release_asset(release, app_name):
    assets = release.get("assets", []) if isinstance(release, dict) else []
    if not assets:
        return None

    app_key = (app_name or "").lower()
    candidates = []
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if not name.endswith(".exe"):
            continue
        if app_key and app_key in name:
            candidates.append(asset)
        elif "installer" in name or "setup" in name:
            candidates.append(asset)

    if candidates:
        return candidates[0]

    for asset in assets:
        if str(asset.get("name", "")).lower().endswith(".exe"):
            return asset
    return None


def get_latest_release(repo_name):
    if not repo_name or repo_name == "YOUR_USERNAME/YOUR_REPO":
        return None

    url = f"https://api.github.com/repos/{repo_name}/releases/latest"
    request = urllib.request.Request(url, headers={"User-Agent": "RaSahLembur-Updater"})
    with urllib.request.urlopen(request, timeout=20) as response:
        data = response.read()
    return json.loads(data)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAutomationUI(root)
    root.mainloop()
