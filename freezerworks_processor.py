import csv
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import webbrowser
import sys
import os
import subprocess
import ssl
import json
from datetime import datetime
from PyPDF2 import PdfMerger
from io import BytesIO
import platform
import time
import traceback


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify=tk.LEFT,
            background="#ffffe0",
            foreground="#000000",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Avenir", 10),
        )
        label.pack(ipadx=6, ipady=4)

    def hide(self, _event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class AliquotUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Freezerworks Processor")
        self.root.geometry("520x620")
        self.root.configure(bg="#f7f7f5")

        style = ttk.Style()
        # Force a light theme so widgets remain readable in macOS dark mode.
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        bg = "#f6f7fb"
        panel = "#f6f7fb"
        text = "#0f172a"
        muted = "#64748b"
        accent = "#2563eb"
        border = "#e2e8f0"
        self._text_color = text
        self._muted_color = muted
        self._token_font_normal = ("Avenir", 12)
        self._token_font_placeholder = ("Avenir", 10)

        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=bg)
        style.configure("TLabel", font=("Avenir", 12), padding=3, background=bg, foreground=text)
        style.configure("Header.TLabel", font=("Avenir", 11, "bold"), padding=3, background=bg, foreground=muted)
        style.configure("TRadiobutton", font=("Avenir", 12, "bold"), background=bg, foreground=text)
        style.configure("TCheckbutton", font=("Avenir", 11), background=bg, foreground=text)
        style.configure("Small.TCheckbutton", font=("Avenir", 10), background=bg, foreground=text)
        style.configure("TEntry", font=("Avenir", 12), padding=3, fieldbackground=bg)
        style.configure("Error.TEntry", font=("Avenir", 12), padding=3, fieldbackground="#fee2e2")
        style.configure("TButton", font=("Avenir", 12), padding=5)
        style.configure("Small.TButton", font=("Avenir", 10), padding=2)
        style.configure("Ghost.TButton", font=("Avenir", 10), padding=2, foreground=muted)
        style.map(
            "TButton",
            background=[("active", "#e8f0fe"), ("!disabled", panel)],
            foreground=[("!disabled", text)],
            bordercolor=[("!disabled", border)],
        )
        style.map(
            "Ghost.TButton",
            background=[("active", "#e8f0fe"), ("!disabled", panel)],
            foreground=[("!disabled", muted)],
            bordercolor=[("!disabled", border)],
        )
        style.configure("Accent.TButton", font=("Avenir", 12, "bold"))
        style.map(
            "Accent.TButton",
            background=[("active", "#e2e8f0"), ("!disabled", panel)],
            foreground=[("!disabled", text)],
        )
        style.configure("TSeparator", background=border)
        style.configure("Status.TLabel", font=("Avenir", 10, "bold"), padding=6, background="#f1f5f9", foreground=text)
        style.configure("Status.Success.TLabel", font=("Avenir", 10, "bold"), padding=6, background="#dcfce7", foreground="#166534")
        style.configure("Status.Error.TLabel", font=("Avenir", 10, "bold"), padding=6, background="#fee2e2", foreground="#991b1b")

        self.main_frame = ttk.Frame(root, padding="12", style="TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        for i in range(0, 10):
            self.main_frame.grid_rowconfigure(i, weight=0)
        self.main_frame.grid_rowconfigure(7, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=0, minsize=190)
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(2, weight=0)

        self.token_label = tk.Label(
            self.main_frame,
            text="Bearer Token*:",
            fg=accent,
            cursor="hand2",
            font=("Avenir", 12, "underline", "bold"),
            bg=bg,
        )
        self.token_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        self.token_label.bind(
            "<Button-1>",
            lambda e: self.open_url(
                "https://freezerworks.pennstatehealth.net/api-docs/elements.html#/"
            ),
        )

        self.token_entry = ttk.Entry(
            self.main_frame, show="*", width=50, cursor="hand2"
        )
        self.token_entry.grid(row=0, column=1, pady=5, sticky=tk.EW)
        self.token_entry.bind("<KeyRelease>", lambda _e: self.update_action_state())
        self.token_entry.bind("<FocusIn>", self._on_token_focus_in)
        self.token_entry.bind("<FocusOut>", self._on_token_focus_out)
        Tooltip(
            self.token_label,
            "Click the label to open the Freezerworks API login and get a token.",
        )
        self._token_placeholder_text = "Paste bearer token here"
        self._token_placeholder_active = False
        self.show_token_var = tk.BooleanVar(value=False)
        self.token_action_frame = ttk.Frame(self.main_frame)
        self.token_action_frame.grid(row=0, column=2, sticky=tk.E, padx=(6, 0))
        self.show_token_checkbox = ttk.Checkbutton(
            self.token_action_frame,
            text="Show",
            variable=self.show_token_var,
            command=self.toggle_token_visibility,
        )
        self.show_token_checkbox.pack(side=tk.LEFT)
        Tooltip(self.show_token_checkbox, "Show or hide the token text.")

        # Checkbox for selecting functionality
        self.functionality_var = tk.StringVar()

        self.process_sample_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Process Patient Sample",
            variable=self.functionality_var,
            value="process_sample",
            command=self.on_workflow_change,
        )
        self.process_sample_checkbox.grid(row=1, column=0, sticky=tk.W, pady=3)
        Tooltip(
            self.process_sample_checkbox,
            "Creates aliquots from a patient sample CSV and prints labels.\n"
            "Required: Sample metadata, dates, and aliquot type.",
        )

        self.download_sample_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_sample_csv,
            style="Ghost.TButton",
        )
        self.download_sample_button.grid(row=1, column=1, sticky=tk.W, pady=3, padx=(12, 0))

        self.freeze_passaged_cells_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Freeze Passaged Cells",
            variable=self.functionality_var,
            value="freeze_passaged_cells",
            command=self.on_workflow_change,
        )

        self.freeze_passaged_cells_checkbox.grid(row=2, column=0, sticky=tk.W, pady=3)
        Tooltip(
            self.freeze_passaged_cells_checkbox,
            "Creates aliquots for passaged cell cultures and prints labels.\n"
            "Required: cell line, dates, passage number, and aliquot type.",
        )

        self.download_freeze_passaged_cells_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_passage_csv,
            style="Ghost.TButton",
        )

        self.download_freeze_passaged_cells_button.grid(
            row=2, column=1, sticky=tk.W, pady=3, padx=(12, 0)
        )

        self.aliquot_assignment_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Aliquot Freezer Assignment",
            variable=self.functionality_var,
            value="aliquot_assignment",
            command=self.on_workflow_change,
        )
        self.aliquot_assignment_checkbox.grid(row=3, column=0, sticky=tk.W, pady=3)
        Tooltip(
            self.aliquot_assignment_checkbox,
            "Updates aliquot freezer locations from a CSV.\n"
            "Required: aliquot ID and location fields.",
        )

        self.download_aliquot_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_aliquot_csv,
            style="Ghost.TButton",
        )
        self.download_aliquot_button.grid(row=3, column=1, sticky=tk.W, pady=3, padx=(12, 0))

        self.file_label = ttk.Label(self.main_frame, text="Select CSV File:", font=("Avenir", 12, "bold"))
        self.file_label.grid(row=4, column=0, sticky=tk.W, pady=4)

        self.file_path = tk.StringVar()
        self.file_path_entry = ttk.Entry(
            self.main_frame,
            textvariable=self.file_path,
            cursor="hand2",
        )
        self.file_path_entry.bind("<Button-1>", lambda e: self.browse_file())
        self.file_path_entry.grid(row=4, column=1, sticky=tk.EW, pady=4)
        Tooltip(self.file_label, "Choose the CSV template for the selected workflow.")
        self.browse_button = ttk.Button(
            self.main_frame, text="Browse...", command=self.browse_file
        )
        self.browse_button.grid(row=4, column=2, sticky=tk.E, padx=(6, 0), pady=4)

        self.dry_run_var = tk.BooleanVar(value=False)
        self.dry_run_checkbox = ttk.Checkbutton(
            self.main_frame,
            text="Dry Run (no API changes)",
            variable=self.dry_run_var,
            command=self.save_settings,
            style="Small.TCheckbutton",
        )
        self.dry_run_checkbox.grid(row=5, column=0, sticky=tk.W, pady=4)
        Tooltip(self.dry_run_checkbox, "Validate inputs without writing to Freezerworks.")

        self.update_button = ttk.Button(
            self.main_frame, text="Update", command=self.start_update, style="Accent.TButton"
        )
        self.update_button.grid(row=5, column=1, columnspan=2, pady=10, sticky=tk.EW)

        self.progress = ttk.Progressbar(self.main_frame, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=4)
        self.progress.grid_remove()

        self.log_frame = ttk.Frame(self.main_frame)
        self.log_frame.grid(row=7, column=0, columnspan=3, pady=6, sticky=tk.NSEW)
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            self.log_frame,
            height=14,
            wrap=tk.WORD,
            font=("Avenir", 10),
            bg=bg,
            fg=text,
            insertbackground=text,
            highlightbackground=border,
            highlightcolor=border,
        )
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)

        self.scrollbar = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.scrollbar.grid(row=0, column=1, sticky=tk.NS, padx=(4, 0))
        self.log_text["yscrollcommand"] = self.scrollbar.set

        self.log_text.tag_configure("bold", font=("Helvetica", 10, "bold"))

        self.log_button_frame = ttk.Frame(self.main_frame)
        self.log_button_frame.grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=(0, 0))
        self.copy_log_button = ttk.Button(
            self.log_button_frame,
            text="Copy Log",
            command=self.copy_log,
            style="Small.TButton",
        )
        self.copy_log_button.pack(side=tk.LEFT)
        self.clear_log_button = ttk.Button(
            self.log_button_frame,
            text="Clear Log",
            command=self.clear_log,
            style="Small.TButton",
        )
        self.clear_log_button.pack(side=tk.LEFT, padx=(6, 0))
        self.open_log_button = ttk.Button(
            self.log_button_frame,
            text="Open Log File",
            command=self.open_log_file,
            style="Small.TButton",
        )
        self.open_log_button.pack(side=tk.LEFT, padx=(6, 0))

        self.set_csv_picker_state(False)
        self.update_action_state()

        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(
            root, textvariable=self.status_var, anchor="w", style="Status.TLabel"
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label.pack_forget()

        # Footer label
        self.footer_label = tk.Label(
            root,
            text="@Thussenthan Walter-Angelo (Sholler Lab, PSCOM); Version 1.7",
            font=("Helvetica", 10),
            bg="#f0f0f0",
            anchor="center",
            fg="#00008B",
            cursor="hand2",
        )
        self.footer_label.pack(side=tk.BOTTOM, fill=tk.X, pady=0)
        self.footer_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://github.com/thussenthan"),
        )

        self.not_updated_aliquots = []
        self.base_url = "https://freezerworks.pennstatehealth.net/api/v1"
        self.cert_path = self.get_cert_path()
        self.log_file_path = os.path.join(self._writable_dir(), "freezerworks_processor.log")
        self.updating = False
        self._ellipse_index = 0
        self.settings_path = os.path.join(self._writable_dir(), "settings.json")
        self.auto_open_file_picker = True
        self.load_settings()
        self._apply_token_placeholder_if_needed()
        # Getting Started popup removed per user preference

    def _resource_dir(self):
        if getattr(sys, "frozen", False):
            return sys._MEIPASS
        return os.path.dirname(__file__)

    def _writable_dir(self):
        if getattr(sys, "frozen", False):
            if platform.system() == "Darwin":
                base = os.path.expanduser("~/Library/Logs/Freezerworks Processor")
                os.makedirs(base, exist_ok=True)
                return base
            return os.path.dirname(sys.executable)
        return os.path.dirname(__file__)

    def _on_ui(self, func, *args, **kwargs):
        if threading.current_thread() is threading.main_thread():
            return func(*args, **kwargs)
        self.root.after(0, lambda: func(*args, **kwargs))

    def show_error(self, title, message):
        self.set_status(f"{title}: {message}")
        self.log(f"{title}: {message}", bold=True)
        self._on_ui(messagebox.showerror, title, message)

    def clear_not_updated_aliquots(self):
        self.not_updated_aliquots.clear()

    def get_cert_path(self):
        resource_dir = self._resource_dir()
        cert_names = [
            "freezerworks.pennstatehealth.net.crt",
            "freezerworks.pennstatehealth.net.cer",
        ]
        search_dirs = [resource_dir, os.getcwd()]

        for directory in search_dirs:
            for name in cert_names:
                candidate = os.path.join(directory, name)
                if os.path.exists(candidate):
                    try:
                        with open(candidate, "rb") as cert_file:
                            data = cert_file.read()
                        if data.lstrip().startswith(b"-----BEGIN CERTIFICATE-----"):
                            return candidate
                        pem_data = ssl.DER_cert_to_PEM_cert(data)
                        pem_path = os.path.join(
                            self._writable_dir(),
                            "freezerworks.pennstatehealth.net.pem",
                        )
                        with open(pem_path, "w") as pem_file:
                            pem_file.write(pem_data)
                        return pem_path
                    except Exception as e:
                        self.show_error(
                            "Error",
                            f"Failed to load SSL certificate: {e}",
                        )
                        return candidate

        searched = "\n".join(f"- {d}" for d in search_dirs)
        names = ", ".join(cert_names)
        self.show_error(
            "Error",
            "SSL Certificate not found.\n"
            f"Looked for {names} in:\n{searched}\n"
            "Place the certificate next to the script/executable or run from that directory.",
        )
        return os.path.join(resource_dir, cert_names[0])

    def browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file:
            self.file_path.set(file)
            self.save_settings()
            self.update_action_state()

    def download_sample_csv(self):
        """Generate and download the default Process Patient Sample CSV template."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="Process Patient Sample.csv",
        )
        if file_path:
            with open(file_path, "w", newline="") as file:
                writer = csv.writer(file)
                # Write header and example row for Process Patient Sample CSV
                writer.writerow(
                    [
                        "Sample Collection Site",
                        "Sample Study ID",
                        "SL0 Number",
                        "Aliquot Type",
                        "Date of Collection",
                        "Freezing Date",
                        "(Study Time Point)",
                        "(Notes)",
                        "(Number of PK Aliquots)",
                        "(PK Time Point)",
                    ]
                )
                today_date = datetime.now().strftime("%m/%d/%Y")
                writer.writerow(
                    [
                        "22",
                        "NMTT-373-03",
                        "3260",
                        "BM, ctDNA, NK, BMA, BC, Tumor, or PK",
                        today_date,
                        today_date,
                        "Day 181",
                    ]
                )

            # set the downloaded file into the file_path entry
            self.file_path.set(file_path)
            self.open_file(file_path)

    def download_passage_csv(self):
        """Generate and download the default Freeze Passaged Cells CSV template."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="Freeze Passaged Cells.csv",
        )
        if file_path:
            with open(file_path, "w", newline="") as file:
                writer = csv.writer(file)
                # Write header and example row for Process Patient Sample CSV
                writer.writerow(
                    [
                        "SL0 Number",
                        "Cell Line Name",
                        "Sample Study ID",
                        "Aliquot Type",
                        "Date of Collection",
                        "Sample Collection Site",
                        "Date of Culture Initiation",
                        "Passage Number",
                        "Freezing Date",
                        "(Media)",
                        "(Serum Supplement)",
                        "(Notes)",
                    ]
                )
                today_date = datetime.now().strftime("%m/%d/%Y")
                writer.writerow(
                    [
                        "3260",
                        "SL03260-191769",
                        "NMTT-373-03",
                        "Bone Core",
                        today_date,
                        "22",
                        today_date,
                        "3",
                        today_date,
                        "RPMI or MEM Alpha or DMEM",
                        "10% FBS",
                    ]
                )

            # set the downloaded file into the file_path entry
            self.file_path.set(file_path)
            self.open_file(file_path)

    def download_aliquot_csv(self):
        """Generate and download the default Aliquot Freezer Assignment CSV template."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="Aliquot Freezer Assignment.csv",
        )
        if file_path:
            with open(file_path, "w", newline="") as file:
                writer = csv.writer(file)
                # Write header and example row for Aliquot Freezer Assignment CSV
                writer.writerow(
                    ["Aliquot ID", "(Shelf)", "Rack", "(Row)", "Box", "Position"]
                )
                writer.writerow(["187093", "", "1", "", "2", "3"])

            # set the downloaded file into the file_path entry
            self.file_path.set(file_path)
            self.open_file(file_path)

    def open_file(self, file_path):
        """Open the file with the default application based on the operating system."""
        if platform.system() == "Windows":
            os.startfile(file_path)  # For Windows
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path], check=False)  # For macOS
        else:
            subprocess.run(["xdg-open", file_path], check=False)  # For Linux

    def run_in_thread(self, target):
        threading.Thread(target=target).start()

    def start_update(self):
        selected_functionality = self.functionality_var.get()
        token = self._get_token_value()
        if not token:
            self._set_token_error(True)
            self.show_error("Error", "Please enter the Bearer Token.")
            return
        if not selected_functionality:
            messagebox.showerror("Error", "Please select a functionality to proceed.")
            return
        csv_file_path = self.file_path.get()
        if csv_file_path:
            self.set_log_file_path(csv_file_path)
        self.save_settings()
        self.set_status("Validating CSV...")
        if not self.validate_csv_schema(selected_functionality, csv_file_path):
            self.set_status("")
            return
        csv_rows = self.read_csv(csv_file_path)
        if not self.preflight_csv_rows(selected_functionality, csv_rows):
            self.set_status("")
            return
        if self.dry_run_var.get():
            self.log("Dry run completed. No API calls were made.", bold=True)
            self.set_status("Dry run complete")
            return
        # begin animation & disable button
        self.updating = True
        self._ellipse_index = 0
        self.update_button.config(state=tk.DISABLED)
        self._set_progress(mode="indeterminate")
        self.animate_update_text()
        self.set_status("Running...")

        if selected_functionality == "aliquot_assignment":
            self.run_in_thread(self._wrapped_update_aliquots)
        elif selected_functionality == "process_sample":
            self.run_in_thread(self._wrapped_process_patient_sample)
        else:  # freeze_passaged_cells
            self.run_in_thread(self._wrapped_passage_culture_cells)

    def _wrapped_update_aliquots(self):
        try:
            self.update_aliquots()
        except Exception as e:
            self.log("Unexpected error while updating aliquots. Check the log file.", bold=True)
            self.log_exception("update_aliquots", e)
        finally:
            self.root.after(0, self.finish_update)

    def _wrapped_process_patient_sample(self):
        try:
            self.process_patient_sample()
        except Exception as e:
            self.log("Unexpected error while processing samples. Check the log file.", bold=True)
            self.log_exception("process_patient_sample", e)
        finally:
            self.root.after(0, self.finish_update)

    def _wrapped_passage_culture_cells(self):
        try:
            self.passage_culture_cells()
        except Exception as e:
            self.log("Unexpected error while freezing passaged cells. Check the log file.", bold=True)
            self.log_exception("passage_culture_cells", e)
        finally:
            self.root.after(0, self.finish_update)

    def animate_update_text(self):
        if not self.updating:
            self.update_button.config(text="Update", state=tk.NORMAL)
            return
        dots = "." * (self._ellipse_index % 4)
        self.update_button.config(text=f"Updating{dots}")
        self._ellipse_index += 1
        # re-schedule
        self.root.after(500, self.animate_update_text)

    def finish_update(self):
        # call this on the main thread when work is done
        self.updating = False
        self._set_progress(mode="determinate", maximum=1, value=0)
        self.progress.grid_remove()
        self.set_status("Completed")
        self.root.after(2000, lambda: self.set_status(""))

    def validate_inputs(self):
        token = self._get_token_value()
        if not token:
            self._set_token_error(True)
            self.show_error("Error", "Please enter the Bearer Token.")
            return None, None
        token = token.strip()
        if token.lower().startswith("bearer "):
            token = token[7:].strip()
        if not token or any(ch.isspace() for ch in token):
            self.show_error(
                "Error",
                "Bearer Token looks invalid (contains whitespace). Please paste the full token only.",
            )
            return None, None

        csv_file_path = self.file_path.get()
        if not csv_file_path:
            self.show_error("Error", "Please select a CSV file.")
            return None, None
        headers = {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + token,
        }

        test_url = f"{self.base_url}/freezers/"
        try:
            test_response = self._request("GET", test_url, headers=headers)
            if test_response.status_code != 200:
                self.log(
                    f"Token validation failed: {test_response.status_code} - {test_response.reason}",
                    bold=True,
                )
                self.show_error(
                    "Error",
                    "Invalid Bearer Token. Please check (or re-generate) your token and try again.",
                )
                return None, None
        except requests.exceptions.SSLError as e:
            self.show_error("SSL Error", f"SSL verification failed: {str(e)}")
            return None, None
        except requests.exceptions.RequestException as e:
            self.show_error("Network Error", f"Connection failed: {str(e)}")
            return None, None

        return headers, csv_file_path

    def read_csv(self, csv_file_path):
        with open(csv_file_path, "r", newline="", encoding="utf-8-sig") as csvfile:
            csv_reader = csv.reader(csvfile)
            try:
                next(csv_reader)  # skip header
            except StopIteration:
                self.show_error("Error", "CSV file is empty.")
                return []
            return list(csv_reader)

    def set_status(self, message):
        def _update():
            if message:
                if not self.status_label.winfo_ismapped():
                    self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
                self.status_var.set(message)
                self.status_label.configure(style=self._status_style_for_message(message))
            else:
                if self.status_label.winfo_ismapped():
                    self.status_label.pack_forget()
                self.status_var.set("")
                self.status_label.configure(style="Status.TLabel")

        self._on_ui(_update)

    def _status_style_for_message(self, message):
        lowered = message.lower()
        if any(word in lowered for word in ("error", "failed", "invalid", "ssl", "network")):
            return "Status.Error.TLabel"
        if any(word in lowered for word in ("completed", "copied", "cleared", "success", "done")):
            return "Status.Success.TLabel"
        return "Status.TLabel"

    def _apply_token_placeholder_if_needed(self):
        if not self.token_entry.get():
            self._set_token_placeholder()

    def _set_token_placeholder(self):
        if self._token_placeholder_active:
            return
        self._token_placeholder_active = True
        self.token_entry.config(
            show="",
            foreground=self._muted_color,
            font=self._token_font_placeholder,
            style="TEntry",
        )
        self.token_entry.delete(0, tk.END)
        self.token_entry.insert(0, self._token_placeholder_text)

    def _clear_token_placeholder(self):
        if not self._token_placeholder_active:
            return
        self._token_placeholder_active = False
        self.token_entry.delete(0, tk.END)
        self.token_entry.config(foreground=self._text_color, font=self._token_font_normal)
        self.toggle_token_visibility()

    def _on_token_focus_in(self, _event=None):
        if self._token_placeholder_active:
            self._clear_token_placeholder()

    def _on_token_focus_out(self, _event=None):
        if not self.token_entry.get().strip():
            self._set_token_placeholder()

    def _get_token_value(self):
        if self._token_placeholder_active:
            return ""
        return self.token_entry.get().strip()

    def _set_token_error(self, has_error):
        style = "Error.TEntry" if has_error else "TEntry"
        self.token_entry.configure(style=style)

    def toggle_token_visibility(self):
        if self._token_placeholder_active:
            self.token_entry.config(show="")
            return
        show_char = "" if self.show_token_var.get() else "*"
        self.token_entry.config(show=show_char)
        self.update_action_state()

    def set_csv_picker_state(self, enabled):
        state = "normal" if enabled else "disabled"
        self.file_path_entry.config(state=state)
        self.browse_button.config(state=state)

    def on_workflow_change(self):
        self.set_csv_picker_state(True)
        self.save_settings()
        self.update_action_state()
        if self.auto_open_file_picker and not self.file_path.get().strip():
            self.root.after(120, self.browse_file)

    def update_action_state(self):
        has_token = bool(self._get_token_value())
        has_workflow = bool(self.functionality_var.get())
        has_csv = bool(self.file_path.get().strip())
        enabled = has_token and has_workflow and has_csv
        if has_token:
            self._set_token_error(False)
        self.update_button.config(state=tk.NORMAL if enabled and not self.updating else tk.DISABLED)


    def _set_progress(self, value=None, maximum=None, mode=None):
        def _update():
            if mode:
                self.progress.config(mode=mode)
                if mode == "indeterminate":
                    if not self.progress.winfo_ismapped():
                        self.progress.grid()
                    self.progress.start(10)
                else:
                    self.progress.stop()
                    if not self.progress.winfo_ismapped():
                        self.progress.grid()
            if maximum is not None:
                self.progress["maximum"] = maximum
            if value is not None:
                self.progress["value"] = value

        self._on_ui(_update)

    def preflight_csv_rows(self, selected_functionality, csv_rows):
        if not csv_rows:
            self.show_error("Error", "CSV file has no data rows.")
            return False
        required_by_workflow = {
            "process_sample": [0, 1, 2, 3, 4, 5],
            "freeze_passaged_cells": [0, 1, 2, 3, 4, 5, 6, 7, 8],
            "aliquot_assignment": [0, 2, 4, 5],
        }
        date_columns = {
            "process_sample": [4, 5],
            "freeze_passaged_cells": [4, 6, 8],
        }
        required_cols = required_by_workflow.get(selected_functionality, [])
        date_cols = date_columns.get(selected_functionality, [])

        total_rows = len(csv_rows)
        self._set_progress(mode="determinate", maximum=total_rows, value=0)
        for idx, row in enumerate(csv_rows, start=2):
            self._set_progress(value=idx - 1)
            if len(row) < max(required_cols, default=0) + 1:
                self.log(f"Error: Insufficient columns in row {idx}: {row}", bold=True)
                self.show_error("CSV Format Error", f"Row {idx} is missing columns.")
                return False
            for col_index in required_cols:
                if not row[col_index].strip():
                    self.log(f"Error: Missing required value in row {idx}.", bold=True)
                    self.show_error(
                        "CSV Format Error",
                        f"Row {idx} is missing required data. Please fix the CSV.",
                    )
                    return False
            for col_index in date_cols:
                value = row[col_index].strip()
                if value:
                    try:
                        datetime.strptime(value, "%m/%d/%Y")
                    except ValueError:
                        self.log(
                            f"Error: Invalid date format in row {idx}: {value}",
                            bold=True,
                        )
                        self.show_error(
                            "CSV Format Error",
                            f"Row {idx} has an invalid date: {value}. Use MM/DD/YYYY.",
                        )
                        return False
        return True

    def save_settings(self):
        try:
            settings = {
                "csv_file_path": self.file_path.get(),
                "workflow": self.functionality_var.get(),
                "dry_run": bool(self.dry_run_var.get()),
            }
            with open(self.settings_path, "w") as settings_file:
                json.dump(settings, settings_file)
        except Exception:
            pass

    def load_settings(self):
        try:
            with open(self.settings_path, "r") as settings_file:
                settings = json.load(settings_file)
            if settings.get("csv_file_path"):
                self.file_path.set(settings["csv_file_path"])
            if settings.get("workflow"):
                self.functionality_var.set(settings["workflow"])
                self.set_csv_picker_state(True)
            if "dry_run" in settings:
                self.dry_run_var.set(bool(settings["dry_run"]))
            self.update_action_state()
        except Exception:
            pass

    def _request(self, method, url, headers=None, **kwargs):
        retries = 3
        backoff = 0.5
        for attempt in range(1, retries + 1):
            try:
                response = requests.request(
                    method,
                    url,
                    headers=headers,
                    verify=self.cert_path,
                    timeout=30,
                    **kwargs,
                )
                if response.status_code in [500, 502, 503, 504] and attempt < retries:
                    self.log(
                        f"Warning: {response.status_code} from {url}. Retrying...",
                        bold=True,
                    )
                    time.sleep(backoff)
                    backoff *= 2
                    continue
                return response
            except requests.exceptions.RequestException as e:
                if attempt >= retries:
                    raise
                self.log(f"Warning: network error {e}. Retrying...", bold=True)
                time.sleep(backoff)
                backoff *= 2

    def _print_labels(self, label_id, aliquot_ids, headers, master_id, aliquot_type=None):
        label_url = f"{self.base_url}/labels/{label_id}/print"
        label_payload = {
            "aliquots": aliquot_ids,
            "numberOfLabelsPerAliquot": 1,
        }
        try:
            response = self._request(
                "POST",
                label_url,
                headers=headers,
                json=label_payload,
            )
            response.raise_for_status()
            if aliquot_type:
                self.log(f"Labels made for SL0 Number {master_id}, {aliquot_type}")
            else:
                self.log(f"Labels made for SL0 Number {master_id}")
            return response.content
        except requests.exceptions.RequestException as e:
            self.log(
                f"Error during label printing for SL0 Number {master_id}: {e}",
                bold=True,
            )
            self.not_updated_aliquots.append(master_id)
            return None

    def validate_csv_schema(self, selected_functionality, csv_file_path):
        if not csv_file_path:
            self.show_error("Error", "Please select a CSV file.")
            return False

        expected_headers = {
            "process_sample": [
                "Sample Collection Site",
                "Sample Study ID",
                "SL0 Number",
                "Aliquot Type",
                "Date of Collection",
                "Freezing Date",
                "(Study Time Point)",
                "(Notes)",
                "(Number of PK Aliquots)",
                "(PK Time Point)",
            ],
            "freeze_passaged_cells": [
                "SL0 Number",
                "Cell Line Name",
                "Sample Study ID",
                "Aliquot Type",
                "Date of Collection",
                "Sample Collection Site",
                "Date of Culture Initiation",
                "Passage Number",
                "Freezing Date",
                "(Media)",
                "(Serum Supplement)",
                "(Notes)",
            ],
            "aliquot_assignment": [
                "Aliquot ID",
                "(Shelf)",
                "Rack",
                "(Row)",
                "Box",
                "Position",
            ],
        }

        expected = expected_headers.get(selected_functionality)
        if not expected:
            self.show_error("Error", "Unknown workflow selected.")
            return False

        try:
            with open(csv_file_path, "r", newline="", encoding="utf-8-sig") as csvfile:
                csv_reader = csv.reader(csvfile)
                header = next(csv_reader, None)
        except OSError as e:
            self.show_error("Error", f"Failed to read CSV file: {e}")
            return False

        if not header:
            self.show_error("Error", "CSV file is empty.")
            return False

        def normalize(value):
            return " ".join(value.strip().split()).lower()

        header_norm = [normalize(h) for h in header]
        expected_norm = [normalize(h) for h in expected]

        missing = [expected[i] for i, v in enumerate(expected_norm) if v not in header_norm]
        extra = [header[i] for i, v in enumerate(header_norm) if v not in expected_norm]

        if header_norm != expected_norm:
            parts = []
            if missing:
                parts.append("Missing columns: " + ", ".join(missing))
            if extra:
                parts.append("Unexpected columns: " + ", ".join(extra))
            if not missing and not extra:
                parts.append("Column order does not match the template.")
            parts.append("Re-download the CSV template and try again.")
            self.log("CSV header validation failed.", bold=True)
            self.log("Header found: " + ", ".join(header), bold=True)
            self.log("Header expected: " + ", ".join(expected), bold=True)
            self.show_error("CSV Format Error", "\n".join(parts))
            return False

        return True

    def convert_date_format(self, date_str, Master_ID):
        try:
            return datetime.strptime(date_str, "%m/%d/%Y").strftime("%Y-%m-%d")
        except ValueError:
            self.log(
                f"Error: Invalid date format for {date_str} for SL0 Number {Master_ID}",
                bold=True,
            )
            return None

    def get_hospital_name(self, hospital_id, Master_ID, headers):
        try:
            normalized_id = int(hospital_id)
        except ValueError:
            self.log(
                f"Error: Invalid hospital ID for SL0 Number {Master_ID}", bold=True
            )
            return None
        hospital_url = f"{self.base_url}/fields/10182"
        try:
            response = self._request("GET", hospital_url, headers=headers)
            response.raise_for_status()
            data = response.json()
            # Extract allowable entries
            allowable_entries = data["properties"]["allowableEntries"]
            # Initialize the lookup table
            lookup_table = {}
            # Process each entry to split the prefix and the description
            for entry in allowable_entries:
                # Split at the first space
                parts = entry.split(" ", 1)
                if len(parts) == 2:
                    key, value = parts
                    # Remove leading zeros from the numeric key and add to the lookup table
                    lookup_table[int(key)] = entry
            hospital_name = lookup_table.get(normalized_id)
        except requests.exceptions.RequestException:
            return None
        return hospital_name

    # Get allowable entries for timepoint
    def allowable_timepoint_entries(self, headers):
        url = f"{self.base_url}/fields/10191"
        response = self._request("GET", url, headers=headers)
        response.raise_for_status()
        data = response.json()
        allowable_entries = data["properties"]["allowableEntries"]
        # Build a dictionary mapping the normalized entry to the original entry.
        normalized_allowable = {
            "".join(entry.lower().split()): entry for entry in allowable_entries
        }
        return normalized_allowable

    def get_allowable_entry(self, value: str, headers):
        normalized_value = "".join(value.lower().split())
        normalized_allowable = self.allowable_timepoint_entries(headers)
        return normalized_allowable.get(normalized_value)

    def studyTimepoint(self, Master_ID, aliquot_id, Study_TimePoint, headers):
        aliquot_url = f"{self.base_url}/aliquots/{aliquot_id}"
        allowable_entry = self.get_allowable_entry(Study_TimePoint, headers)

        if allowable_entry is None:
            aliquot_payload = {
                "Study_timepoint_other": Study_TimePoint,
                "pk_time_point": "Other",
            }
        else:
            aliquot_payload = {"pk_time_point": Study_TimePoint}
        time.sleep(0.5)  # Adding a small delay to avoid overwhelming the server
        try:
            response = self._request(
                "POST",
                aliquot_url,
                headers=headers,
                json=aliquot_payload,
            )
            response.raise_for_status()
            return
        except requests.exceptions.RequestException as e:
            self.log(f"Study timepoint request failed: {e} for SL0 Number {Master_ID}")
            self.not_updated_aliquots.append(Master_ID)

    def masterID_search(self, Master_ID, headers):
        aliquot_url = ""
        # Payload for initial sample search
        if len(Master_ID) == 3:
            Master_ID = "0" + Master_ID
        payload = {
            "table": "Samples",
            "listViewId": -9,
            "limit": 1,
            "searchLines": [
                {
                    "lineNumber": 1,
                    "fieldId": 10166,
                    "comparison": "in",
                    "compareValue": f"{Master_ID}",
                    "openParensCount": 0,
                    "closeParensCount": 0,
                }
            ],
        }

        try:
            response = self._request(
                "POST",
                f"{self.base_url}/search/",
                headers=headers,
                json=payload,
            )
            response.raise_for_status()
            data = response.json()
        except requests.exceptions.RequestException as e:
            self.log(
                f"Error during search request for SL0 Number {Master_ID}: {e}",
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        try:
            freezerworks_id = data["properties"]["results"][0]["FreezerworksID"]
        except (KeyError, IndexError):
            self.log(
                f"Error: Unable to find FreezerworksID in response data for SL0 Number {Master_ID}",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        aliquot_url = f"{self.base_url}/samples/{freezerworks_id}/aliquots"
        return aliquot_url

    def output_merged_pdf(self, merger, files_added):
        # After all processing is done, output the merged PDF
        if files_added > 0:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile="Processed_Samples.pdf",
            )
            if output_path:
                # Write the merged PDF to file
                with open(output_path, "wb") as output_file:
                    merger.write(output_file)
                    self.log(f"PDF successfully saved to {output_path}", bold=True)
                self.open_file(output_path)
        else:
            self.log("No labels to print.")
        merger.close()

        if self.not_updated_aliquots:
            self.log("The following SL0 Numbers were not updated correctly:", bold=True)
            for master_ID in self.not_updated_aliquots:
                self.log(master_ID, bold=True)
        else:
            self.log("All SL0 Numbers were updated successfully.", bold=True)

    def process_patient_sample(self):
        headers, csv_file_path = self.validate_inputs()
        if headers is None or csv_file_path is None:
            return

        self.clear_not_updated_aliquots()

        csv_rows = self.read_csv(csv_file_path)

        merger = PdfMerger()
        files_added = 0

        total_rows = len(csv_rows)
        self._set_progress(mode="determinate", maximum=total_rows, value=0)
        for index, csv_row in enumerate(csv_rows, start=1):
            self._set_progress(value=index)
            self.log(f"Processing row {index}/{total_rows}")
            if len(csv_row) < 6:
                self.not_updated_aliquots.append("Incorrect CSV Formatting")
                self.log(f"Error: Insufficient columns in row: {csv_row}", bold=True)
                continue
            (
                Sample_Collection_Site,
                Sample_Study_ID,
                Master_ID,
                Aliquot_Type,
                Date_of_Collection,
                Freezing_Date,
                Study_TimePoint,
                Notes,
                Number_of_PK_Aliquots,
                PK_Time_Point,
            ) = [col.strip() for col in csv_row[:10]]

            if Master_ID:  # Only process if Master ID is present
                sample_pdf = self.process_sample(
                    headers,
                    Master_ID,
                    Aliquot_Type,
                    Date_of_Collection,
                    Sample_Collection_Site,
                    Freezing_Date,
                    Sample_Study_ID,
                    Study_TimePoint,
                    Notes,
                    Number_of_PK_Aliquots,
                    PK_Time_Point,
                )
                if sample_pdf:
                    pdf_stream = BytesIO(sample_pdf)  # Use BytesIO for PDF stream
                    try:
                        merger.append(pdf_stream)  # Append to the merger
                        files_added += 1  # Increment files added count
                    except Exception as e:
                        self.log(
                            f"Error appending PDF labels for SL0 Number {Master_ID}: {e}",
                            bold=True,
                        )
        self.output_merged_pdf(merger, files_added)  # Output the merged PDF

    def process_sample(
        self,
        headers,
        Master_ID,
        Aliquot_Type,
        Date_of_Collection,
        Sample_Collection_Site,
        Freezing_Date,
        Sample_Study_ID,
        Study_TimePoint,
        Notes,
        Number_of_PK_Aliquots,
        PK_Time_Point,
    ):
        if (
            not Master_ID
            or not Aliquot_Type
            or not Date_of_Collection
            or not Freezing_Date
            or not Sample_Study_ID
            or not Sample_Collection_Site
        ):
            self.log(
                f"Error: Missing required fields for SL0 Number {Master_ID}", bold=True
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        # Get hospital name and check if it's valid
        Hospital_Name = self.get_hospital_name(
            Sample_Collection_Site, Master_ID, headers
        )
        if Hospital_Name is None:
            self.log(
                f"Error: Sample processing aborted due to invalid hospital name for SL0 Number {Master_ID}.",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        # Convert dates to YYYY-MM-DD format
        Date_of_Collection = self.convert_date_format(Date_of_Collection, Master_ID)
        Freezing_Date = self.convert_date_format(Freezing_Date, Master_ID)
        if Date_of_Collection is None or Freezing_Date is None:
            self.log(
                f"Error: Sample processing aborted due to invalid date format for SL0 Number {Master_ID}",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        # Initialize list to store aliquot IDs for label printing
        labels_to_print_ids = []

        aliquot_url = self.masterID_search(Master_ID, headers)

        aliquot_payload = {
            "numberOfAliquots": 1,
            "WorkflowStatus": "Available",
            "Date_of_Collection": Date_of_Collection,
            "Sample_Collection_Site": Hospital_Name,
            "Sample_Notes": Notes,
            "Sample_Study_ID": Sample_Study_ID,
        }

        if Aliquot_Type == "PK":
            if Number_of_PK_Aliquots.strip():
                Number_of_PK_Aliquots = int(Number_of_PK_Aliquots)
            else:
                self.log(
                    f"Error: Number of PK Aliquots is required for SL0 Number {Master_ID}",
                    bold=True,
                )
                self.not_updated_aliquots.append(Master_ID)
                return
            aliquot_payload.update(
                {
                    "numberOfAliquots": Number_of_PK_Aliquots,
                    "Aliquot_Type": "Plasma for PK analysis",
                    "Freezing_Date": Freezing_Date,
                    "PK_Time_Point_W_R_To_Dose": PK_Time_Point,
                }
            )

            # Make aliquot creation requests
            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                for entity in data["entities"]:
                    labels_to_print_ids.append(entity["PK_AliquotUID"])
                    self.log(f"Aliquot UID created: {entity['PK_AliquotUID']}")
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            if Study_TimePoint:
                for aliquot_id in labels_to_print_ids:
                    self.studyTimepoint(Master_ID, aliquot_id, Study_TimePoint, headers)

            # Printing labels
            label_id = 17 if "BCC18" in Sample_Study_ID and Aliquot_Type == "ADA Serum" else 4
            return self._print_labels(label_id, labels_to_print_ids, headers, Master_ID, Aliquot_Type)
        elif Aliquot_Type == "BMA":
            aliquot_payload.update({"Aliquot_Type": "Bone Marrow Aspirate"})

            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                PK_ParentAliquotID = data["properties"]["PK_AliquotUID"]
                self.log(f"Aliquot UID created: {PK_ParentAliquotID}")
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            aliquot_payload.update(
                {
                    "Subaliquot_Type": "MNC",
                    "Freezing_Date": Freezing_Date,
                    "PK_ParentAliquotID": PK_ParentAliquotID,
                }
            )

            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                pk_aliquot_uid = data["properties"]["PK_AliquotUID"]
                labels_to_print_ids.append(pk_aliquot_uid)
                self.log(f"Aliquot UID created: {pk_aliquot_uid}")

                aliquot_url_BMA = f"{self.base_url}/aliquots/{PK_ParentAliquotID}"
                aliquot_payload_BMA = {
                    "Passage_number": 0,
                    "Cell_Line_Name_": f"SL0{Master_ID}-{PK_ParentAliquotID}",
                    "Subaliquot_Type": "Cultured",
                }
                time.sleep(0.5)
                response = self._request(
                    "POST",
                    aliquot_url_BMA,
                    headers=headers,
                    json=aliquot_payload_BMA,
                )
                response.raise_for_status()
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            if Study_TimePoint:
                self.studyTimepoint(
                    Master_ID, PK_ParentAliquotID, Study_TimePoint, headers
                )
                for aliquot_id in labels_to_print_ids:
                    self.studyTimepoint(Master_ID, aliquot_id, Study_TimePoint, headers)

            return self._print_labels(9, labels_to_print_ids, headers, Master_ID, Aliquot_Type)

        elif Aliquot_Type in ["BC", "Tumor"]:
            if Aliquot_Type == "BC":
                Aliquot_Type = "Bone Core"
            aliquot_payload.update(
                {
                    "Aliquot_Type": Aliquot_Type,
                    "Subaliquot_Type": "Cultured",
                    "Passage_number": 0,
                }
            )

            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                PK_ParentAliquotID = data["properties"]["PK_AliquotUID"]
                labels_to_print_ids.append(PK_ParentAliquotID)
                self.log(f"Aliquot UID created: {PK_ParentAliquotID}")

                time.sleep(0.5)  # Adding a small delay to avoid overwhelming the server

                aliquot_url_alt = f"{self.base_url}/aliquots/{PK_ParentAliquotID}"
                aliquot_payload_alt = {
                    "Cell_Line_Name_": f"SL0{Master_ID}-{PK_ParentAliquotID}"
                }
                response = self._request(
                    "POST",
                    aliquot_url_alt,
                    headers=headers,
                    json=aliquot_payload_alt,
                )
                response.raise_for_status()
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            if Study_TimePoint:
                self.studyTimepoint(
                    Master_ID, PK_ParentAliquotID, Study_TimePoint, headers
                )
                for aliquot_id in labels_to_print_ids:
                    self.studyTimepoint(Master_ID, aliquot_id, Study_TimePoint, headers)

            return self._print_labels(3, labels_to_print_ids, headers, Master_ID, Aliquot_Type)

        else:
            if Aliquot_Type == "BM":
                Aliquot_Type = "Biomarker Blood"

            if Aliquot_Type == "NK":
                Aliquot_Type = "NK cell analysis"

            aliquot_payload.update(
                {
                    "Aliquot_Type": Aliquot_Type,
                }
            )

            # Make aliquot creation requests
            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                PK_ParentAliquotID = data["properties"]["PK_AliquotUID"]
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            # Determine the number of repeats and subaliquot type
            if Aliquot_Type in ["ctDNA", "Biomarker Blood"]:
                repeats = 4
            elif Aliquot_Type == "NK cell analysis":
                repeats = 2
            else:
                self.log(
                    f"Error: Incorrect Aliquot_Type Data Format: {Aliquot_Type} for SL0 Number {Master_ID}",
                    bold=True,
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            aliquot_payload.update(
                {
                    "Subaliquot_Type": "Plasma",
                    "Freezing_Date": Freezing_Date,
                    "PK_ParentAliquotID": PK_ParentAliquotID,
                }
            )

            # Loop through for repeats
            for _ in range(repeats):
                try:
                    response = self._request(
                        "POST",
                        aliquot_url,
                        headers=headers,
                        json=aliquot_payload,
                    )
                    response.raise_for_status()
                    data = response.json()
                    pk_aliquot_uid = data["properties"]["PK_AliquotUID"]
                    labels_to_print_ids.append(pk_aliquot_uid)
                    self.log(f"Aliquot UID created: {pk_aliquot_uid}")
                except requests.exceptions.RequestException as e:
                    self.log(
                        f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                    )
                    self.not_updated_aliquots.append(Master_ID)
                    return

            if Aliquot_Type in ["ctDNA", "Biomarker Blood"]:
                Subaliquot_Type = "Buffy Coat"
            elif Aliquot_Type == "NK cell analysis":
                Subaliquot_Type = "MNC"
            else:
                self.log(
                    f"Error: Incorrect Aliquot_Type Data Format: {Aliquot_Type} for SL0 Number {Master_ID}",
                    bold=True,
                )
                self.not_updated_aliquots.append(Master_ID)
                return

            aliquot_payload.update(
                {
                    "Subaliquot_Type": Subaliquot_Type,
                }
            )

            for _ in range(2):
                try:
                    response = self._request(
                        "POST",
                        aliquot_url,
                        headers=headers,
                        json=aliquot_payload,
                    )
                    response.raise_for_status()
                    data = response.json()
                    pk_aliquot_uid = data["properties"]["PK_AliquotUID"]
                    labels_to_print_ids.append(pk_aliquot_uid)
                    self.log(f"Aliquot UID created: {pk_aliquot_uid}")
                except requests.exceptions.RequestException as e:
                    self.log(
                        f"Error during aliquot creation (in repeat loop) for SL0 Number {Master_ID}: {e}",
                    )
                    return
            if Study_TimePoint:
                self.studyTimepoint(
                    Master_ID, PK_ParentAliquotID, Study_TimePoint, headers
                )
                for aliquot_id in labels_to_print_ids:
                    self.studyTimepoint(Master_ID, aliquot_id, Study_TimePoint, headers)

            return self._print_labels(9, labels_to_print_ids, headers, Master_ID, Aliquot_Type)

    def passage_culture_cells(self):
        headers, csv_file_path = self.validate_inputs()
        if headers is None or csv_file_path is None:
            return

        self.clear_not_updated_aliquots()

        csv_rows = self.read_csv(csv_file_path)

        merger = PdfMerger()
        files_added = 0

        total_rows = len(csv_rows)
        self._set_progress(mode="determinate", maximum=total_rows, value=0)
        for index, csv_row in enumerate(csv_rows, start=1):
            self._set_progress(value=index)
            self.log(f"Processing row {index}/{total_rows}")
            if len(csv_row) < 11:
                self.not_updated_aliquots.append("Incorrect CSV Formatting")
                self.log(f"Error: Insufficient columns in row: {csv_row}", bold=True)
                continue
            (
                Master_ID,
                Cell_Line_Name,
                Sample_Study_ID,
                Aliquot_Type,
                Date_of_Collection,
                Sample_Collection_Site,
                Date_of_Culture_Initiation,
                Passage_Number,
                Freezing_Date,
                Media,
                Serum_Supplement,
                Notes,
            ) = [col.strip() for col in csv_row[:12]]

            if Master_ID:  # Only process if Master ID is present
                sample_pdf = self.passage_cells(
                    headers,
                    Master_ID,
                    Cell_Line_Name,
                    Sample_Study_ID,
                    Aliquot_Type,
                    Date_of_Collection,
                    Sample_Collection_Site,
                    Date_of_Culture_Initiation,
                    Passage_Number,
                    Freezing_Date,
                    Media,
                    Serum_Supplement,
                    Notes,
                )

                if sample_pdf:
                    pdf_stream = BytesIO(sample_pdf)  # Use BytesIO for PDF stream
                    merger.append(pdf_stream)  # Append to the merger
                    files_added += 1

        self.output_merged_pdf(merger, files_added)  # Output the merged PDF

    def passage_cells(
        self,
        headers,
        Master_ID,
        Cell_Line_Name,
        Sample_Study_ID,
        Aliquot_Type,
        Date_of_Collection,
        Sample_Collection_Site,
        Date_of_Culture_Initiation,
        Passage_Number,
        Freezing_Date,
        Media,
        Serum_Supplement,
        Notes,
    ):
        if (
            not Master_ID
            or not Cell_Line_Name
            or not Sample_Study_ID
            or not Aliquot_Type
            or not Date_of_Collection
            or not Sample_Collection_Site
            or not Date_of_Culture_Initiation
            or not Passage_Number
            or not Freezing_Date
        ):
            self.log(
                f"Error: Missing required fields for SL0 Number {Master_ID}", bold=True
            )
            self.not_updated_aliquots.append(Master_ID)
            return
        # Get hospital name and check if it's valid
        Hospital_Name = self.get_hospital_name(
            Sample_Collection_Site, Master_ID, headers
        )
        if Hospital_Name is None:
            self.log(
                f"Error: Sample processing aborted due to invalid hospital name for SL0 Number {Master_ID}.",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        # Convert dates to YYYY-MM-DD format
        Date_of_Collection = self.convert_date_format(Date_of_Collection, Master_ID)
        Freezing_Date = self.convert_date_format(Freezing_Date, Master_ID)
        Date_of_Culture_Initiation = self.convert_date_format(
            Date_of_Culture_Initiation, Master_ID
        )
        if (
            Date_of_Collection is None
            or Freezing_Date is None
            or Date_of_Culture_Initiation is None
        ):
            self.log(
                f"Error: Sample processing aborted due to invalid date format for SL0 Number {Master_ID}",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        IDEXX_labels_to_print_ids = []
        frozen_culture_labels_to_print_ids = []

        aliquot_url = self.masterID_search(Master_ID, headers)

        # Determine the number of repeats and label types
        Passage_Number = int(Passage_Number)
        if Passage_Number == 1:
            labels_ids = [3]
        elif Passage_Number == 2:
            labels_ids = [3, 3, 3]
        elif Passage_Number == 3:
            labels_ids = [3, 3, 3, 3, 7]
        elif Passage_Number in [
            4,
            5,
            6,
            7,
            8,
            9,
        ]:
            labels_ids = [3, 3, 3, 3]
        else:
            self.log(
                f"Error: Incorrect Passage_Number Data Format for SL0 Number {Master_ID}",
                bold=True,
            )
            self.not_updated_aliquots.append(Master_ID)
            return

        if Passage_Number in [2, 4]:
            self.log(
                f"Remember to Image Cells for SL0 Number {Master_ID}!",
                bold=True,
            )

        if Serum_Supplement == "10% FBS":
            Serum_Supplement = "10%  FBS"
        if Serum_Supplement == "20% FBS":
            Serum_Supplement = "20%  FBS"

        aliquot_payload = {
            "numberOfAliquots": 1,
            "WorkflowStatus": "Available",
            "Subaliquot_Type": "Cultured",
            "Cell_Line_Name_": Cell_Line_Name,
            "Sample_Study_ID": Sample_Study_ID,
            "Aliquot_Type": Aliquot_Type,
            "Date_of_Collection": Date_of_Collection,
            "Sample_Collection_Site": Hospital_Name,
            "Date_of_Culture_Initiation": Date_of_Culture_Initiation,
            "Passage_number": Passage_Number,
            "Freezing_Date": Freezing_Date,
            "Media": Media,
            "Serum_Supplement": Serum_Supplement,
            "Sample_Notes": Notes,
        }

        # Loop through for repeats
        for label_id in labels_ids:
            try:
                response = self._request(
                    "POST",
                    aliquot_url,
                    headers=headers,
                    json=aliquot_payload,
                )
                response.raise_for_status()
                data = response.json()
                pk_aliquot_uid = data["properties"]["PK_AliquotUID"]

                if label_id == 3:
                    frozen_culture_labels_to_print_ids.append(pk_aliquot_uid)
                elif label_id == 7:
                    IDEXX_labels_to_print_ids.append(pk_aliquot_uid)
                    time.sleep(0.5)
                    IDEXX_string = "IDEXX " + Notes
                    IDEXX_aliquot_payload = {"Sample_Notes": IDEXX_string}
                    try:
                        response = self._request(
                            "POST",
                            f"{self.base_url}/aliquots/{pk_aliquot_uid}",
                            headers=headers,
                            json=IDEXX_aliquot_payload,
                        )
                    except requests.exceptions.RequestException as e:
                        self.log(
                            f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                        )
                        self.not_updated_aliquots.append(Master_ID)
                        return
                self.log(f"Aliquot UID created: {pk_aliquot_uid}")
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during aliquot creation for SL0 Number {Master_ID}: {e}",
                )
                self.not_updated_aliquots.append(Master_ID)
                return

        # Printing labels
        if Passage_Number == 3:
            pdf_merger = PdfMerger()

            frozen_content = self._print_labels(
                3, frozen_culture_labels_to_print_ids, headers, Master_ID
            )
            if not frozen_content:
                return
            frozen_culture_pdf = BytesIO(frozen_content)
            pdf_merger.append(frozen_culture_pdf)

            idexx_content = self._print_labels(
                7, IDEXX_labels_to_print_ids, headers, Master_ID
            )
            if not idexx_content:
                return
            IDEXX_pdf = BytesIO(idexx_content)
            pdf_merger.append(IDEXX_pdf)

            merged_label_pdf = BytesIO()
            pdf_merger.write(merged_label_pdf)
            pdf_merger.close()
            merged_label_pdf.seek(0)
            self.log(f"Labels made for SL0 Number {Master_ID}")
            return merged_label_pdf.getvalue()

        else:
            return self._print_labels(
                3, frozen_culture_labels_to_print_ids, headers, Master_ID
            )

    def update_aliquots(self):
        headers, csv_file_path = self.validate_inputs()
        if headers is None or csv_file_path is None:
            return

        self.clear_not_updated_aliquots()

        csv_rows = self.read_csv(csv_file_path)

        total_rows = len(csv_rows)
        self._set_progress(mode="determinate", maximum=total_rows, value=0)
        for index, csv_row in enumerate(csv_rows, start=1):
            self._set_progress(value=index)
            self.log(f"Processing row {index}/{total_rows}")
            if len(csv_row) < 4:
                self.not_updated_aliquots.append("Incorrect CSV Formatting")
                self.log(f"Error: Insufficient columns in row: {csv_row}", bold=True)
                continue

            aliquot_id, shelf, rack, row, box, position = [
                col.strip() for col in csv_row[:6]
            ]

            if aliquot_id:  # Only process if aliquot ID is present
                self.update_aliquot(
                    headers, aliquot_id, shelf, rack, row, box, position
                )

        if self.not_updated_aliquots:
            self.log("The following aliquots were not updated correctly:", bold=True)
            for aliquot in self.not_updated_aliquots:
                self.log(aliquot, bold=True)
        else:
            self.log("All aliquots were updated successfully.", bold=True)

    def update_aliquot(self, headers, aliquot_id, shelf, rack, row, box, position):
        url = f"{self.base_url}/aliquots/{aliquot_id}"

        try:
            output_str = self._request("GET", url, headers=headers)
            output_str.raise_for_status()
            try:
                output = output_str.json()
            except ValueError:
                self.log(
                    f"Error: Failed to parse response for aliquot {aliquot_id}: {output_str.status_code}",
                    bold=True,
                )
                self.not_updated_aliquots.append(aliquot_id)
                return

            position1_value = output["properties"].get("position1")

            if not position1_value or position1_value == "null":
                if (not shelf and not row) and all(
                    s.isdigit() for s in [rack, box, position]
                ):
                    payload = {
                        "fwLocation": {
                            "FK_FreezerSectID": 10019,
                            "position1": rack,
                            "position2": box,
                            "position3": position,
                        }
                    }
                elif all(s.isdigit() for s in [shelf, rack, row, box, position]):
                    payload = {
                        "fwLocation": {
                            "FK_FreezerSectID": 10014,
                            "position1": shelf,
                            "position2": rack,
                            "position3": row,
                            "position4": box,
                            "position5": position,
                        }
                    }
                else:  # Handle invalid data here
                    self.not_updated_aliquots.append(aliquot_id)
                    self.log(
                        f"Error: error in .csv spreadsheet file for aliquot row {aliquot_id}",
                        bold=True,
                    )
                    return
                response = self._request(
                    "POST", url, headers=headers, json=payload
                )

                # Handle response status code
                if response.status_code != 200:
                    self.log(
                        f"Error: Failed to update aliquot {aliquot_id}: {response.status_code} - {response.reason}",
                        bold=True,
                    )
                    self.not_updated_aliquots.append(aliquot_id)
                else:
                    self.log(f"Updated aliquot {aliquot_id}")
            else:
                self.not_updated_aliquots.append(aliquot_id)
                self.log(
                    f"Error: aliquot already assigned location {aliquot_id}",
                    bold=True,
                )
        except requests.exceptions.RequestException as e:
            self.log(f"SSL Error for aliquot {aliquot_id}: {str(e)}", bold=True)
            self.not_updated_aliquots.append(aliquot_id)

    def log(self, message, bold=False):
        self.append_log_file(message)
        tag = "bold" if bold else None
        timestamp = datetime.now().strftime("%H:%M:%S")
        ui_message = f"[{timestamp}] {message}"
        self._on_ui(self._append_log, ui_message, tag)

    def _append_log(self, message, tag):
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)

    def log_exception(self, context, exc):
        try:
            with open(self.log_file_path, "a") as log_file:
                log_file.write(
                    f"\n[{datetime.now().isoformat()}] {context}: {repr(exc)}\n"
                )
                log_file.write(traceback.format_exc())
        except Exception:
            # Avoid crashing on logging failures.
            pass

    def open_log_file(self):
        self.open_file(self.log_file_path)

    def copy_log(self):
        try:
            content = self.log_text.get("1.0", tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.set_status("Log copied to clipboard")
        except Exception:
            self.set_status("Failed to copy log")

    def clear_log(self):
        try:
            self.log_text.delete("1.0", tk.END)
            with open(self.log_file_path, "w") as log_file:
                log_file.write("")
            self.set_status("Log cleared")
        except Exception:
            self.set_status("Failed to clear log")

    def set_log_file_path(self, csv_file_path):
        try:
            csv_dir = os.path.dirname(os.path.abspath(csv_file_path))
            self.log_file_path = os.path.join(csv_dir, "freezerworks_processor.log")
        except Exception:
            self.log_file_path = os.path.join(
                self._writable_dir(), "freezerworks_processor.log"
            )

    def append_log_file(self, message):
        try:
            with open(self.log_file_path, "a") as log_file:
                log_file.write(f"[{datetime.now().isoformat()}] {message}\n")
        except Exception:
            # Avoid crashing on logging failures.
            pass

    def open_url(self, url):
        webbrowser.open_new(url)


if __name__ == "__main__":
    root = tk.Tk()
    app = AliquotUpdaterApp(root)
    root.mainloop()
