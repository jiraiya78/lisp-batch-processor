#!/usr/bin/env python
"""
Revision: 9 - Updated status messages for each Lisp execution with the Lisp name.
Created by Jiraiya78 | Version 1.0.3

Changes:
- After executing each Lisp script for a DWG file, the status is updated to indicate
  which Lisp (by name and order) completed for that file.
  For example: "myscript.lsp completed for file X (Lisp 1 of 3)"
- Other functionalities remain unchanged.
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import pythoncom
import win32com.client
import win32gui
import win32con
import threading
import json
from PIL import Image, ImageTk
import time

def resource_path(relative_path):
    """
    Get absolute path to resource, works for PyInstaller.
    If running as a bundled executable, sys._MEIPASS points to the temporary folder
    where resources are stored. Otherwise, return the path relative to current directory.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Set the TKDND_LIBRARY environment variable for tkinterDnD.
os.environ["TKDND_LIBRARY"] = resource_path("tkdnd2.8")

def hide_autocad_window():
    """
    Enumerate all top-level windows and hide those whose title contains 'AutoCAD'.
    """
    def enum_handler(hwnd, lparam):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "AutoCAD" in title:
                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    win32gui.EnumWindows(enum_handler, None)

def get_lisp_files(directory):
    """Recursively scan for .lsp files in a given directory."""
    lisp_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".lsp"):
                lisp_files.append(os.path.join(root, file))
    return lisp_files

class LispBatchProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lisp Batch Processor")
        self.root.resizable(False, False)
        self.file_list = []
        # Change lisp_files to a list of dicts with keys "path" and "var" for maintaining order.
        self.lisp_files = []

        # Determine the base path for the application.
        if getattr(sys, "frozen", False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        self.default_lisp_dir = os.path.join(base_path, "lisp")

        self.settings_file = "settings.json"
        self.load_settings()
        self.load_default_lisps()
        self.create_widgets()
        self.style_widgets()
        self.update_process_button_state()
        self.options_window = None
        self.success_count = 0

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, "r") as f:
                self.settings = json.load(f)
        else:
            self.settings = {"autocad_location": ""}
            self.save_settings()

    def save_settings(self):
        with open(self.settings_file, "w") as f:
            json.dump(self.settings, f, indent=4)

    def load_default_lisps(self):
        # Scan the default Lisp directory (and subdirectories) for .lsp files.
        if os.path.isdir(self.default_lisp_dir):
            default_lisps = get_lisp_files(self.default_lisp_dir)
            for lisp in default_lisps:
                self.lisp_files.append({"path": lisp, "var": tk.BooleanVar(value=True)})
        else:
            os.makedirs(self.default_lisp_dir, exist_ok=True)

    def create_widgets(self):
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        file_frame = tk.LabelFrame(frame, text="DWG Files", padx=10, pady=10)
        file_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        lisp_frame = tk.LabelFrame(frame, text="Lisp Scripts", padx=10, pady=10)
        lisp_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.file_listbox = tk.Listbox(file_frame, selectmode=tk.EXTENDED, width=50, height=15, font=("Helvetica", 12), activestyle="none")
        self.file_listbox.grid(row=0, column=0, sticky="nsew")

        file_scrollbar = tk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 5))
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)

        self.file_listbox.bind("<Delete>", lambda e: self.remove_files())
        self.file_listbox.bind("<<ListboxSelect>>", self.update_backdrop_text)

        self.backdrop_text = tk.Label(self.file_listbox, text="Drag and drop to add file or use button", font=("Helvetica", 12, "italic"), fg="grey")
        self.backdrop_text.pack(side="top", fill="both", expand=True)

        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.grid(row=0, column=2, padx=(5, 0), pady=5, sticky="n")

        self.add_file_button = tk.Button(file_buttons_frame, text="+", command=self.add_files, font=("Helvetica", 24, "bold"), fg="green", width=2, height=1)
        self.add_file_button.pack(pady=(10, 5))

        self.remove_file_button = tk.Button(file_buttons_frame, text="-", command=self.remove_files, font=("Helvetica", 24, "bold"), fg="red", width=2, height=1)
        self.remove_file_button.pack(pady=(5, 10))

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_files)

        # Create a frame to hold the list of Lisp entries with reordering buttons.
        self.lisp_listbox_frame = tk.Frame(lisp_frame)
        self.lisp_listbox_frame.pack(pady=5, fill=tk.BOTH, expand=True)
        self.refresh_lisp_list()

        lisp_buttons_frame = tk.Frame(lisp_frame)
        lisp_buttons_frame.pack(pady=5)
        self.add_lisp_button = ttk.Button(lisp_buttons_frame, text="Add Lisp", command=self.add_lisp)
        self.add_lisp_button.pack(side=tk.LEFT, padx=5)
        self.remove_lisp_button = ttk.Button(lisp_buttons_frame, text="Remove Lisp", command=self.remove_lisp)
        self.remove_lisp_button.pack(side=tk.LEFT, padx=5)

        self.process_button = ttk.Button(self.root, text="Process Files", command=self.start_processing, width=20)
        self.process_button.pack(pady=10)
        self.process_button.config(state=tk.DISABLED)

        self.progress = ttk.Progressbar(self.root, length=600, mode="determinate")
        self.progress.pack(pady=5, fill=tk.X, padx=10)

        status_frame = tk.Frame(self.root)
        status_frame.pack(pady=5, fill=tk.X, padx=10)
        self.status_text = tk.Text(status_frame, height=7, font=("Helvetica", 12, "italic"), state="disabled")
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        status_scrollbar = tk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=status_scrollbar.set)

        # Use resource_path to correctly load the gear image.
        gear_img_path = resource_path("images/gear.png")
        try:
            gear_img =  Image.open(gear_img_path)
            gear_img = gear_img.resize((24, 24), Image.Resampling.LANCZOS)
            self.gear_photo = ImageTk.PhotoImage(gear_img)
        except Exception as e:
            print("Failed to load gear image:", e)
            self.gear_photo = None
        self.options_button = ttk.Button(self.root, image=self.gear_photo, command=self.open_options)
        self.options_button.pack(side=tk.RIGHT, padx=10, pady=10)
        self.credit_label = ttk.Label(self.root, text="Created by Jiraiya78 | Version 1.0.3", font=("Helvetica", 10, "italic"))
        self.credit_label.pack(side=tk.BOTTOM, pady=5)

    def refresh_lisp_list(self):
        """Clear and redraw the list of Lisp scripts with up/down buttons."""
        # Remove existing widgets in lisp_listbox_frame.
        for widget in self.lisp_listbox_frame.winfo_children():
            widget.destroy()
        # For each Lisp entry, create a frame with a checkbutton and arrow buttons.
        for index, item in enumerate(self.lisp_files):
            row_frame = tk.Frame(self.lisp_listbox_frame)
            row_frame.pack(fill=tk.X, pady=2)

            chk = tk.Checkbutton(row_frame, text=os.path.basename(item["path"]), variable=item["var"], font=("Helvetica", 12))
            chk.pack(side=tk.LEFT, padx=5)
            
            btn_frame = tk.Frame(row_frame)
            btn_frame.pack(side=tk.RIGHT)
            
            # Up button (disable if first item)
            up_state = tk.NORMAL if index > 0 else tk.DISABLED
            up_btn = ttk.Button(btn_frame, text="▲", width=2, state=up_state, command=lambda idx=index: self.move_lisp_up(idx))
            up_btn.pack(side=tk.LEFT, padx=2)
            # Down button (disable if last item)
            down_state = tk.NORMAL if index < len(self.lisp_files) - 1 else tk.DISABLED
            down_btn = ttk.Button(btn_frame, text="▼", width=2, state=down_state, command=lambda idx=index: self.move_lisp_down(idx))
            down_btn.pack(side=tk.LEFT, padx=2)

    def move_lisp_up(self, index):
        """Move the Lisp script at the given index up in the list."""
        if index > 0:
            self.lisp_files[index - 1], self.lisp_files[index] = self.lisp_files[index], self.lisp_files[index - 1]
            self.refresh_lisp_list()

    def move_lisp_down(self, index):
        """Move the Lisp script at the given index down in the list."""
        if index < len(self.lisp_files) - 1:
            self.lisp_files[index + 1], self.lisp_files[index] = self.lisp_files[index], self.lisp_files[index + 1]
            self.refresh_lisp_list()

    def style_widgets(self):
        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 12, "bold"), padding=6)
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TListbox", font=("Courier", 12))
        style.configure("TProgressbar", thickness=20)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("DWG Files", "*.dwg")])
        for file in files:
            if file not in self.file_list:
                self.file_list.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        self.update_process_button_state()
        self.update_backdrop_text()

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.endswith(".dwg") and file not in self.file_list:
                self.file_list.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
        self.update_process_button_state()
        self.update_backdrop_text()

    def remove_files(self):
        selected_files = self.file_listbox.curselection()
        for index in reversed(selected_files):
            self.file_listbox.delete(index)
            del self.file_list[index]
        self.update_process_button_state()
        self.update_backdrop_text()

    def update_backdrop_text(self, event=None):
        if not self.file_list:
            self.backdrop_text.pack(side="top", fill="both", expand=True)
        else:
            self.backdrop_text.pack_forget()

    def add_lisp(self):
        lisp_files = filedialog.askopenfilenames(filetypes=[("Lisp Files", "*.lsp")])
        for lisp in lisp_files:
            lisp_path = os.path.abspath(lisp)
            # Only add new Lisp if not already in the list.
            if not any(item["path"] == lisp_path for item in self.lisp_files):
                self.lisp_files.append({"path": lisp_path, "var": tk.BooleanVar(value=True)})
        self.refresh_lisp_list()
        self.update_process_button_state()

    def remove_lisp(self):
        # Remove entries where the checkbutton is unchecked.
        self.lisp_files = [item for item in self.lisp_files if item["var"].get()]
        self.refresh_lisp_list()
        self.update_process_button_state()

    def update_process_button_state(self):
        if self.file_list and any(item["var"].get() for item in self.lisp_files):
            self.process_button.config(state=tk.NORMAL)
        else:
            self.process_button.config(state=tk.DISABLED)

    def start_processing(self):
        self.disable_buttons()
        self.success_count = 0
        threading.Thread(target=self.process_files).start()

    def process_files(self):
        pythoncom.CoInitialize()
        total_files = len(self.file_list)
        self.update_status("Initializing AutoCAD...", "blue")
        try:
            acad_location = self.settings["autocad_location"]
            if not os.path.exists(acad_location):
                self.update_status(f"AutoCAD.exe not found at {acad_location}", "red")
                self.enable_buttons()
                return

            acad = win32com.client.Dispatch("AutoCAD.Application")
            acad.Visible = False
            acad.WindowState = 1
            hide_autocad_window()

            # Use the order defined in lisp_files.
            selected_lisps = [item["path"] for item in self.lisp_files if item["var"].get()]

            for index, file in enumerate(self.file_list):
                self.update_status(f"Processing file: {os.path.basename(file)} ({index+1}/{total_files})", "blue")
                self.update_progress(index+1, total_files)
                try:
                    self.run_lisp_process(acad, file, selected_lisps)
                    self.update_status(f"Process successful for file {file}", "green")
                    self.success_count += 1
                except Exception as e:
                    error_str = str(e)
                    if "Open.Close" in error_str:
                        self.update_status(f"Error processing file {file}: The file could not be opened or closed.", "red")
                    elif "disconnected" in error_str:
                        self.update_status(f"Error processing file {file}: AutoCAD may have crashed.", "red")
                    else:
                        self.update_status(f"Error processing file {file}: {e}", "red")

            try:
                acad.Quit()
            except Exception as quit_exception:
                self.updateStatus(f"Error quitting AutoCAD: {quit_exception}", "red")
        except Exception as e:
            self.update_status(f"Error initializing AutoCAD: {e}", "red")
        finally:
            self.update_status(f"Processing complete: {self.success_count} of {total_files} files processed successfully.", "blue")
            self.update_progress(total_files, total_files)
            pythoncom.CoUninitialize()
            self.enable_buttons()

    def run_lisp_process(self, acad, file, selected_lisps):
        try:
            doc = self.safe_open_document(acad, file)
            if doc:
                for i, lisp in enumerate(selected_lisps):
                    lisp_path_fixed = lisp.replace("\\", "/")
                    self.send_command_with_retry(acad, f'(load "{lisp_path_fixed}")\n')
                    time.sleep(1)
                    self.send_command_with_retry(acad, f'(c:MyLispFunction)\n')
                    time.sleep(1)
                    # Extract the Lisp name for the status message.
                    lisp_name = os.path.basename(lisp)
                    self.update_status(f'{lisp_name} completed for file {os.path.basename(file)} (Lisp {i+1} of {len(selected_lisps)})', "blue")
                self.send_command_with_retry(acad, '(command "_.QSAVE")\n')
                time.sleep(2)
                self.send_command_with_retry(acad, '(command "_.CLOSE")\n')
                time.sleep(3)
                if self.is_document_open(acad, file):
                    self.update_status(f"Warning: Document did not close properly on first attempt for {file}", "orange")
                    self.send_command_with_retry(acad, '(command "_.CLOSE")\n')
                    time.sleep(3)
                if not self.is_document_open(acad, file):
                    try:
                        doc.Close(SaveChanges=True)
                    except Exception as close_exception:
                        self.update_status(f"Suppressed final close error for {file}: {close_exception}", "orange")
                else:
                    self.update_status(f"Warning: Document still appears open for {file}", "orange")
            else:
                raise Exception("Failed to open document after multiple attempts")
        except Exception as e:
            raise e

    def is_document_open(self, acad, file_path):
        try:
            for doc in acad.Documents:
                if os.path.normcase(doc.FullName) == os.path.normcase(file_path):
                    return True
            return False
        except Exception:
            return False

    def safe_open_document(self, acad, file, retries=5, delay=4):
        for attempt in range(retries):
            try:
                return acad.Documents.Open(file)
            except Exception as e:
                if attempt < retries - 1:
                    self.update_status(f"Retrying to open file {file}... (Attempt {attempt+1}/{retries})", "orange")
                    time.sleep(delay)
                else:
                    raise e

    def send_command_with_retry(self, acad, command, retries=3):
        for attempt in range(retries):
            try:
                acad.ActiveDocument.SendCommand(command)
                time.sleep(1)
                return
            except Exception as e:
                if attempt < retries - 1:
                    time.sleep(2)
                else:
                    raise e

    def update_status(self, status, color="blue"):
        self.root.after(0, self._set_status_text, status, color)

    def _set_status_text(self, status, color):
        self.status_text.config(state="normal")
        self.status_text.insert("end", f"{status}\n")
        if color == "red":
            self.status_text.tag_configure("error", foreground="red")
            self.status_text.tag_add("error", "end-2l", "end-1c")
        elif color == "green":
            self.status_text.tag_configure("success", foreground="green")
            self.status_text.tag_add("success", "end-2l", "end-1c")
        elif color == "orange":
            self.status_text.tag_configure("warning", foreground="orange")
            self.status_text.tag_add("warning", "end-2l", "end-1c")
        else:
            self.status_text.tag_configure("info", foreground="blue")
            self.status_text.tag_add("info", "end-2l", "end-1c")
        self.status_text.config(state="disabled")
        self.status_text.see("end")

    def update_progress(self, current, total):
        progress_value = (current / total) * 100
        self.root.after(0, self._set_progress, progress_value)

    def _set_progress(self, value):
        self.progress["value"] = value

    def disable_buttons(self):
        self.root.after(0, lambda: self._set_buttons_state(tk.DISABLED))

    def enable_buttons(self):
        self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))

    def _set_buttons_state(self, state):
        self.add_file_button.config(state=state)
        self.remove_file_button.config(state=state)
        self.add_lisp_button.config(state=state)
        self.remove_lisp_button.config(state=state)
        self.process_button.config(state=state)
        self.options_button.config(state=state)

    def open_options(self):
        if self.options_window and self.options_window.winfo_exists():
            return
        self.options_window = tk.Toplevel(self.root)
        self.options_window.title("Options")
        self.options_window.geometry("400x200")
        self.options_window.transient(self.root)
        self.options_window.grab_set()
        self.root.update_idletasks()
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        self.options_window.geometry(f"+{x + width//2 - 200}+{y + height//2 - 100}")
        options_frame = tk.Frame(self.options_window, padx=10, pady=10)
        options_frame.pack(fill=tk.BOTH, expand=True)
        autocad_label = ttk.Label(options_frame, text="AutoCAD Location:", font=("Helvetica", 12))
        autocad_label.pack(anchor="w", pady=5)
        self.autocad_entry = ttk.Entry(options_frame, font=("Helvetica", 12))
        self.autocad_entry.pack(fill=tk.X, pady=5, padx=10)
        if self.settings["autocad_location"]:
            self.autocad_entry.insert(0, self.settings["autocad_location"])
        else:
            self.autocad_entry.insert(0, self.find_autocad_location())
        autocad_browse_button = ttk.Button(options_frame, text="Browse...", command=self.browse_autocad)
        autocad_browse_button.pack(pady=5)
        save_button = ttk.Button(options_frame, text="Save", command=self.save_options)
        save_button.pack(pady=10)

    def find_autocad_location(self):
        possible_paths = [
            "C:\\Program Files\\Autodesk",
            "C:\\Program Files (x86)\\Autodesk"
        ]
        for path in possible_paths:
            for root_dir, dirs, files in os.walk(path):
                if "acad.exe" in files:
                    return os.path.join(root_dir, "acad.exe")
        return ""

    def browse_autocad(self):
        initial_dir = os.path.dirname(self.settings.get("autocad_location", ""))
        if not initial_dir:
            initial_dir = "C:\\Program Files\\Autodesk\\"
        filepath = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("AutoCAD Executable", "acad.exe")])
        if filepath:
            self.autocad_entry.delete(0, tk.END)
            self.autocad_entry.insert(0, filepath)

    def save_options(self):
        autocad_location = self.autocad_entry.get()
        if os.path.basename(autocad_location) == "acad.exe" and os.path.exists(autocad_location):
            self.settings["autocad_location"] = autocad_location
            self.save_settings()
            messagebox.showinfo("Settings Saved", "AutoCAD location has been updated.")
        else:
            messagebox.showerror("Invalid Path", "The specified AutoCAD location is invalid.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = LispBatchProcessorApp(root)
    root.mainloop()