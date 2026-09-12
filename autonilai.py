import time
import pyautogui
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import random

# -------- BASIC CONFIG --------
WAIT_SECONDS = 3
TYPE_DELAY = 0.02
DEFAULT_LAST_COLUMN = ""
# ------------------------------

class ExcelAutomationUI:
    def __init__(self, root):
        self.root = root
        self.root.title("input nilai bot")
        self.root.geometry("560x650")
        self.root.resizable(True, True)
        
        self.wb = None
        self.sheet = None
        self.start_row = 1
        self.excel_path = tk.StringVar()
        self.current_sheet = tk.StringVar()
        self.randomize_delay = tk.BooleanVar(value=False)
        self.min_delay = tk.DoubleVar(value=0.02)
        self.max_delay = tk.DoubleVar(value=0.2)
        self.end_column = tk.StringVar(value=DEFAULT_LAST_COLUMN)
        self.end_row = tk.IntVar(value=1)
        
        # Automation state control
        self.automation_running = False
        self.pause_event = threading.Event()
        self.stop_event = threading.Event()
        self.pause_event.set()  # Not paused initially
        self.resume_countdown = tk.IntVar(value=0)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=0)
        
        # ===== FILE SELECTION =====
        file_label = ttk.Label(main_frame, text="Excel File:", font=("Arial", 11, "bold"))
        file_label.grid(row=0, column=0, columnspan=1, sticky=tk.W, pady=(0, 5))
        
        file_entry = ttk.Entry(main_frame, textvariable=self.excel_path, width=50)
        file_entry.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=1, column=1, sticky=tk.E)
        
        # ===== SHEET SELECTION =====
        sheet_label = ttk.Label(main_frame, text="Select Sheet:", font=("Arial", 11, "bold"))
        sheet_label.grid(row=2, column=0, columnspan=2, sticky=tk.W)
        
        self.sheet_dropdown = ttk.Combobox(main_frame, textvariable=self.current_sheet, state="readonly", width=58)
        self.sheet_dropdown.grid(row=3, column=0, columnspan=1, sticky=(tk.W, tk.E))
        self.sheet_dropdown.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_selected())
        
        # ===== START/END ROW =====
        row_label = ttk.Label(main_frame, text="Starting Row:", font=("Arial", 11, "bold"))
        row_label.grid(row=4, column=0, columnspan=2, sticky=tk.W)
        
        self.row_spinbox = ttk.Spinbox(main_frame, from_=1, to=10000, width=12)
        self.row_spinbox.set(1)
        self.row_spinbox.grid(row=5, column=0, columnspan=2, sticky=tk.W)
        
        end_row_label = ttk.Label(main_frame, text="Last Row:", font=("Arial", 11, "bold"))
        end_row_label.grid(row=6, column=0, columnspan=2, sticky=tk.W)
        
        self.end_row_spinbox = ttk.Spinbox(main_frame, from_=1, to=10000, width=12, textvariable=self.end_row)
        self.end_row_spinbox.set(1)
        self.end_row_spinbox.grid(row=7, column=0, columnspan=2, sticky=tk.W)
        
        # ===== START/END COLUMN =====
        col_label = ttk.Label(main_frame, text="Starting Column:", font=("Arial", 11, "bold"))
        col_label.grid(row=8, column=0, columnspan=2, sticky=tk.W)
        
        self.col_entry = ttk.Entry(main_frame, width=12)
        self.col_entry.grid(row=9, column=0, columnspan=2, sticky=tk.W)
        
        end_col_label = ttk.Label(main_frame, text="Last Column:", font=("Arial", 11, "bold"))
        end_col_label.grid(row=10, column=0, columnspan=2, sticky=tk.W)
        
        self.end_col_entry = ttk.Entry(main_frame, width=12, textvariable=self.end_column)
        self.end_col_entry.grid(row=11, column=0, columnspan=2, sticky=tk.W)
        
        # ===== RANDOMIZE DELAY =====
        randomize_label = ttk.Label(main_frame, text="Randomize Delays:", font=("Arial", 11, "bold"))
        randomize_label.grid(row=12, column=0, columnspan=1, sticky=tk.W)
        
        self.randomize_check = ttk.Checkbutton(main_frame, text="Enable random key press delays", variable=self.randomize_delay)
        self.randomize_check.grid(row=13, column=0, sticky=tk.W)
        
        delay_range_label = ttk.Label(main_frame, text="Delay Range (seconds):", font=("Arial", 10))
        delay_range_label.grid(row=15, column=0, sticky=tk.W, pady=(0, 5))
        
        delay_frame = ttk.Frame(main_frame)
        delay_frame.grid(row=16, column=0, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(delay_frame, text="Min:").pack(side=tk.LEFT, padx=(0, 5))
        self.min_delay_spinbox = ttk.Spinbox(delay_frame, from_=0.01, to=5.0, width=8, textvariable=self.min_delay, format="%.2f")
        self.min_delay_spinbox.pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Label(delay_frame, text="Max:").pack(side=tk.LEFT, padx=(0, 5))
        self.max_delay_spinbox = ttk.Spinbox(delay_frame, from_=0.01, to=5.0, width=8, textvariable=self.max_delay, format="%.2f")
        self.max_delay_spinbox.pack(side=tk.LEFT)
        
        # ===== DATA PREVIEW =====
        preview_label = ttk.Label(main_frame, text="Data Preview:", font=("Arial", 11, "bold"))
        preview_label.grid(row=17, column=0, sticky=tk.W, pady=(15, 5))
        
        # Listbox with scrollbar
        frame_listbox = ttk.Frame(main_frame)
        frame_listbox.grid(row=18, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(frame_listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_listbox = tk.Listbox(frame_listbox, height=8, yscrollcommand=scrollbar.set)
        self.preview_listbox.pack(side=tk.LEFT, fill=(tk.BOTH), expand=True)
        scrollbar.config(command=self.preview_listbox.yview)
        
        # ===== STATUS =====
        self.status_label = ttk.Label(main_frame, text="Status: Ready", foreground="blue")
        self.status_label.grid(row=19, column=0, sticky=tk.W, pady=(10, 0))
        
        # ===== BUTTONS =====
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=20, column=0, sticky=(tk.W, tk.E), pady=(15, 0))
        
        self.execute_btn = ttk.Button(button_frame, text="Execute Automation", command=self.execute_automation)
        self.execute_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.preview_btn = ttk.Button(button_frame, text="Preview Data", command=self.preview_data)
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_automation, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.reset_btn = ttk.Button(button_frame, text="Reset", command=self.reset_automation)
        self.reset_btn.pack(side=tk.LEFT)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if file_path:
            self.excel_path.set(file_path)
            self.load_workbook(file_path)
            
    def load_workbook(self, path):
        try:
            self.wb = load_workbook(path, data_only=True)
            self.sheet_dropdown['values'] = self.wb.sheetnames
            if self.wb.sheetnames:
                self.sheet_dropdown.current(0)
                self.on_sheet_selected()
            self.update_status(f"Loaded: {os.path.basename(path)}", "green")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {e}")
            self.update_status("Error loading file", "red")
            
    def on_sheet_selected(self):
        if self.wb:
            sheet_name = self.current_sheet.get()
            self.sheet = self.wb[sheet_name]
            self.preview_listbox.delete(0, tk.END)
            
    def get_columns(self, start_column, end_column=None):
        if end_column is None:
            end_column = self.end_column.get()
        start_idx = column_index_from_string(start_column)
        end_idx = column_index_from_string(end_column)
        if end_idx < start_idx:
            raise ValueError("Last column must be the same or to the right of the starting column.")
        return [get_column_letter(i) for i in range(start_idx, end_idx + 1)]

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
                cell_value = self.sheet[f"{col}{row}"].value
                values.append("" if cell_value is None else str(cell_value))
            rows.append(values)
        return rows

    def preview_data(self):
        if not self.sheet:
            messagebox.showwarning("Warning", "Please load an Excel file and select a sheet first.")
            return
            
        col = self.col_entry.get().strip().upper()
        if not col:
            messagebox.showwarning("Warning", "Please enter a starting column letter.")
            return
            
        end_col = self.end_column.get().strip().upper()
        if not end_col:
            messagebox.showwarning("Warning", "Please enter a last column letter.")
            return
            
        try:
            start_row = int(self.row_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid starting row.")
            return
            
        try:
            end_row = int(self.end_row_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid last row.")
            return
            
        try:
            rows = self.read_rows(col, start_row, end_col, end_row)
        except Exception as e:
            messagebox.showerror("Error", f"Invalid column or sheet data: {e}")
            return
            
        self.preview_listbox.delete(0, tk.END)
        for row_values in rows:
            self.preview_listbox.insert(tk.END, " | ".join(row_values))
            
        if not rows:
            messagebox.showinfo("Info", f"No data found from {col}{start_row} to {end_col}{end_row}.")
            self.update_status("No data found", "orange")
        else:
            self.update_status(f"Preview: {len(rows)} row(s) from {col} to {end_col}", "blue")
            
    def execute_automation(self):
        if not self.sheet:
            messagebox.showwarning("Warning", "Please load an Excel file and select a sheet first.")
            return
            
        col = self.col_entry.get().strip().upper()
        if not col:
            messagebox.showwarning("Warning", "Please enter a column letter.")
            return
            
        end_col = self.end_column.get().strip().upper()
        if not end_col:
            messagebox.showwarning("Warning", "Please enter a last column letter.")
            return
            
        try:
            start_row = int(self.row_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid starting row.")
            return
            
        try:
            end_row = int(self.end_row_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid last row.")
            return
            
        if end_row < start_row:
            messagebox.showerror("Error", "Last row must be the same or greater than starting row.")
            return
            
        try:
            self.get_columns(col, end_col)
        except Exception as e:
            messagebox.showerror("Error", f"Invalid column range: {e}")
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
                self.update_status("Paused", "orange")
            else:
                # Resume
                self.pause_event.set()
                self.pause_btn.config(text="Pause")
                self.resume_countdown.set(3)  # 3 second countdown
                self.update_status("Resuming in 3s...", "blue")
    
    def stop_automation(self):
        self.stop_event.set()
        self.pause_event.set()  # Ensure not paused so it can exit
        self.update_status("Stopping...", "orange")
    
    def reset_automation(self):
        self.stop_event.set()
        self.pause_event.set()
        self.resume_countdown.set(0)
        self.automation_running = False
        self.execute_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.preview_listbox.delete(0, tk.END)
        self.update_status("Reset", "blue")
        
    def run_automation(self, col, start_row, end_row):
        try:
            # Reset countdown at start
            self.resume_countdown.set(0)
            
            # Read row-by-row data from start to end column and row
            end_col = self.end_column.get().strip().upper()
            rows = self.read_rows(col, start_row, end_col, end_row)
            if not rows:
                self.update_status("No data found", "orange")
                self.automation_running = False
                self.execute_btn.config(state=tk.NORMAL)
                self.pause_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
            self.update_status(f"Starting automation: {len(rows)} row(s) from {col} to {end_col}...", "blue")
            
            # Wait for user to focus target app
            for remaining in range(WAIT_SECONDS, 0, -1):
                if self.stop_event.is_set():
                    self.update_status("Stopped", "red")
                    return
                self.update_status(f"Click target app... ({remaining}s)", "orange")
                time.sleep(1)
                
            # Type into target app row by row
            self.update_status("Automating...", "blue")
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
                            self.update_status(f"Resuming... ({remaining}s)", "orange")
                            time.sleep(1)
                        self.resume_countdown.set(0)
                    
                    #pyautogui.hotkey("ctrl", "a")
                    try:
                        root.clipboard_clear()
                        root.clipboard_append(str(text))
                        root.update()
                        time.sleep(0.05)
                        pyautogui.hotkey("ctrl", "v")
                    except Exception:
                        pyautogui.write(str(text), interval=TYPE_DELAY)
                    pyautogui.press("down")
                    cell_count += 1
                    
                    # Delay after each cell if enabled
                    if self.randomize_delay.get():
                        delay = random.uniform(self.min_delay.get(), self.max_delay.get())
                    else:
                        delay = 0.2
                    time.sleep(delay)
                    self.update_status(f"Automated {cell_count}/{total_cells} cells", "blue")
                
            self.update_status("Automation complete!", "green")
        except Exception as e:
            self.update_status(f"Error: {e}", "red")
        finally:
            self.automation_running = False
            self.resume_countdown.set(0)
            self.execute_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            
    def update_status(self, msg, color="blue"):
        self.status_label.config(text=f"Status: {msg}", foreground=color)
        self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAutomationUI(root)
    root.mainloop()
