import csv
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import webbrowser
import sys
import os
from datetime import datetime
from PyPDF2 import PdfMerger
from io import BytesIO
import platform
import time


class AliquotUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Freezerworks Processor")
        self.root.geometry("600x600")
        self.root.configure(bg="#f0f0f0")

        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 12), padding=5)
        style.configure("TButton", font=("Helvetica", 12), padding=5)
        style.configure("TEntry", font=("Helvetica", 12), padding=5)
        style.configure("Small.TButton", font=("Helvetica", 10), padding=3)

        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)
        self.main_frame.grid_rowconfigure(4, weight=1)
        self.main_frame.grid_rowconfigure(5, weight=1)
        self.main_frame.grid_rowconfigure(6, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(2, weight=1)

        self.token_label = tk.Label(
            self.main_frame,
            text="Enter your Bearer Token:",
            fg="blue",
            cursor="hand2",
            font=("Helvetica", 12, "underline"),
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

        # Checkbox for selecting functionality
        self.functionality_var = tk.StringVar()

        self.process_sample_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Process Patient Sample",
            variable=self.functionality_var,
            value="process_sample",
        )
        self.process_sample_checkbox.grid(row=1, column=0, sticky=tk.W, pady=5)

        self.download_sample_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_sample_csv,
            style="Small.TButton",
        )
        self.download_sample_button.grid(row=1, column=1, sticky=tk.W, pady=5)

        self.freeze_passaged_cells_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Freeze Passaged Cells",
            variable=self.functionality_var,
            value="freeze_passaged_cells",
        )

        self.freeze_passaged_cells_checkbox.grid(row=2, column=0, sticky=tk.W, pady=5)

        self.download_freeze_passaged_cells_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_passage_csv,
            style="Small.TButton",
        )

        self.download_freeze_passaged_cells_button.grid(
            row=2, column=1, sticky=tk.W, pady=5
        )

        self.aliquot_assignment_checkbox = ttk.Radiobutton(
            self.main_frame,
            text="Aliquot Freezer Assignment",
            variable=self.functionality_var,
            value="aliquot_assignment",
        )
        self.aliquot_assignment_checkbox.grid(row=3, column=0, sticky=tk.W, pady=5)

        self.download_aliquot_button = ttk.Button(
            self.main_frame,
            text="Download CSV",
            command=self.download_aliquot_csv,
            style="Small.TButton",
        )
        self.download_aliquot_button.grid(row=3, column=1, sticky=tk.W, pady=5)

        self.file_label = ttk.Label(self.main_frame, text="Select CSV File:")
        self.file_label.grid(row=4, column=0, sticky=tk.W, pady=5)

        self.file_path = tk.StringVar()
        self.file_path_entry = ttk.Entry(
            self.main_frame,
            textvariable=self.file_path,
            cursor="hand2",
        )
        self.file_path_entry.bind("<Button-1>", lambda e: self.browse_file())
        self.file_path_entry.grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=5)

        self.update_button = ttk.Button(
            self.main_frame, text="Update", command=self.start_update
        )
        self.update_button.grid(row=5, column=0, columnspan=2, pady=15, sticky=tk.EW)

        self.log_text = tk.Text(
            self.main_frame,
            height=10,
            wrap=tk.WORD,
            font=("Helvetica", 10),
            bg="#f5f5f5",
        )
        self.log_text.grid(row=6, column=0, columnspan=2, pady=10, sticky=tk.NSEW)

        self.scrollbar = ttk.Scrollbar(self.main_frame, command=self.log_text.yview)
        self.scrollbar.grid(row=6, column=2, sticky=tk.NS)
        self.log_text["yscrollcommand"] = self.scrollbar.set

        self.log_text.tag_configure("bold", font=("Helvetica", 10, "bold"))

        # Footer label
        self.footer_label = tk.Label(
            root,
            text="@Thussenthan Walter-Angelo (Sholler Lab, Penn State College of Medicine); Version 1.6",
            font=("Helvetica", 10),
            bg="#f0f0f0",
            anchor="center",
            fg="#00008B",
            cursor="hand2",
        )
        self.footer_label.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        self.footer_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://github.com/thussenthan"),
        )

        self.not_updated_aliquots = []
        self.base_url = "https://freezerworks.pennstatehealth.net/api/v1"
        self.cert_path = self.get_cert_path()
        self.updating = False
        self._ellipse_index = 0

    def clear_not_updated_aliquots(self):
        self.not_updated_aliquots.clear()

    def get_cert_path(self):
        if getattr(sys, "frozen", False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(__file__)
        cert_path = os.path.join(
            application_path, "freezerworks.pennstatehealth.net.crt"
        )

        if not os.path.exists(cert_path):
            messagebox.showerror(
                "Error", "SSL Certificate not found. Please check the path."
            )

        return cert_path

    def browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file:
            self.file_path.set(file)

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
            os.system(f"open {file_path}")  # For macOS
        else:
            os.system(f"xdg-open {file_path}")  # For Linux

    def run_in_thread(self, target):
        threading.Thread(target=target).start()

    def start_update(self):
        selected_functionality = self.functionality_var.get()
        if not selected_functionality:
            messagebox.showerror("Error", "Please select a functionality to proceed.")
            return
        # begin animation & disable button
        self.updating = True
        self._ellipse_index = 0
        self.update_button.config(state=tk.DISABLED)
        self.animate_update_text()

        if selected_functionality == "aliquot_assignment":
            self.run_in_thread(self._wrapped_update_aliquots)
        elif selected_functionality == "process_sample":
            self.run_in_thread(self._wrapped_process_patient_sample)
        else:  # freeze_passaged_cells
            self.run_in_thread(self._wrapped_passage_culture_cells)

    def _wrapped_update_aliquots(self):
        try:
            self.update_aliquots()
        finally:
            self.root.after(0, self.finish_update)

    def _wrapped_process_patient_sample(self):
        try:
            self.process_patient_sample()
        finally:
            self.root.after(0, self.finish_update)

    def _wrapped_passage_culture_cells(self):
        try:
            self.passage_culture_cells()
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

    def validate_inputs(self):
        token = self.token_entry.get()
        if not token:
            messagebox.showerror("Error", "Please enter the Bearer Token.")
            return None, None

        csv_file_path = self.file_path.get()
        if not csv_file_path:
            messagebox.showerror("Error", "Please select a CSV file.")
            return None, None

        headers = {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + token,
        }

        test_url = f"{self.base_url}/freezers/"
        try:
            test_response = requests.get(
                test_url, headers=headers, verify=self.cert_path
            )
            if test_response.status_code != 200:
                messagebox.showerror(
                    "Error",
                    "Invalid Bearer Token. Please check (or re-generate) your token and try again.",
                )
                return
        except requests.exceptions.SSLError as e:
            messagebox.showerror("SSL Error", f"SSL verification failed: {str(e)}")
            return

        return headers, csv_file_path

    def read_csv(self, csv_file_path):
        with open(csv_file_path, "r") as csvfile:
            csv_reader = csv.reader(csvfile)
            next(csv_reader)
            return list(csv_reader)

    def convert_date_format(self, date_str, Master_ID):
        try:
            return datetime.strptime(date_str, "%m/%d/%Y").strftime("%Y-%m-%d")
        except ValueError:
            self.log(
                f"Error: Invalid date format for {date_str} for SL0 Number {Master_ID}",
                bold=True,
            )
            return None

    def get_hospital_name(self, hospital_id, Master_ID):
        headers, _ = self.validate_inputs()
        if headers is not None:
            try:
                normalized_id = int(hospital_id)
            except ValueError:
                self.log(
                    f"Error: Invalid hospital ID for SL0 Number {Master_ID}", bold=True
                )
                return None
            hospital_url = f"{self.base_url}/fields/10182"
            try:
                response = requests.get(
                    hospital_url,
                    headers=headers,
                    verify=self.cert_path,
                )
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
            except requests.exceptions.RequestException as e:
                return None
            return hospital_name

    # Get allowable entries for timepoint
    def allowable_timepoint_entries(self, headers):
        url = f"{self.base_url}/fields/10191"
        response = requests.get(url, headers=headers, verify=self.cert_path)
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

        if allowable_entry == None:
            aliquot_payload = {
                "Study_timepoint_other": Study_TimePoint,
                "pk_time_point": "Other",
            }
        else:
            aliquot_payload = {"pk_time_point": Study_TimePoint}
        time.sleep(0.5)  # Adding a small delay to avoid overwhelming the server
        try:
            response = requests.post(
                aliquot_url,
                json=aliquot_payload,
                headers=headers,
                verify=self.cert_path,
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
            response = requests.post(
                f"{self.base_url}/search/",
                json=payload,
                headers=headers,
                verify=self.cert_path,
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

        for csv_row in csv_rows:
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
        Hospital_Name = self.get_hospital_name(Sample_Collection_Site, Master_ID)
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

        def print_labels(self, Master_ID, labels_to_print_ids, headers):
            # Printing labels
            label_url = f"{self.base_url}/labels/9/print"
            label_payload = {
                "aliquots": labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    label_url,
                    json=label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                self.log(f"Labels made for SL0 Number {Master_ID}, {Aliquot_Type}")
                return response.content
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label printing for SL0 Number {Master_ID}: {e}",
                    bold=True,
                )
                self.not_updated_aliquots.append(Master_ID)
                return

        def print_cell_line_labels(self, Master_ID, labels_to_print_ids, headers):
            # Printing labels for cell line cultures
            culture_label_url = f"{self.base_url}/labels/3/print"
            culture_label_payload = {
                "aliquots": labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    culture_label_url,
                    json=culture_label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                self.log(f"Labels made for SL0 Number {Master_ID}, {Aliquot_Type}")
                return response.content
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label printing for SL0 Number {Master_ID}: {e}",
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
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
            if "BCC18" in Sample_Study_ID and Aliquot_Type == "ADA Serum":
                label_url = f"{self.base_url}/labels/17/print"
            else:
                label_url = f"{self.base_url}/labels/4/print"
            label_payload = {
                "aliquots": labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    label_url,
                    json=label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                self.log(f"Labels made for SL0 Number {Master_ID}, {Aliquot_Type}")
                return response.content
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label printing for SL0 Number {Master_ID}: {e}",
                )
                return
        elif Aliquot_Type == "BMA":
            aliquot_payload.update({"Aliquot_Type": "Bone Marrow Aspirate"})

            try:
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
                response = requests.post(
                    aliquot_url_BMA,
                    json=aliquot_payload_BMA,
                    headers=headers,
                    verify=self.cert_path,
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

            content = print_labels(self, Master_ID, labels_to_print_ids, headers)
            return content

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
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
                response = requests.post(
                    aliquot_url_alt,
                    json=aliquot_payload_alt,
                    headers=headers,
                    verify=self.cert_path,
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

            content = print_cell_line_labels(
                self, Master_ID, labels_to_print_ids, headers
            )
            return content

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
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
                    response = requests.post(
                        aliquot_url,
                        json=aliquot_payload,
                        headers=headers,
                        verify=self.cert_path,
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
                    response = requests.post(
                        aliquot_url,
                        json=aliquot_payload,
                        headers=headers,
                        verify=self.cert_path,
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

            content = print_labels(self, Master_ID, labels_to_print_ids, headers)
            return content

    def passage_culture_cells(self):
        headers, csv_file_path = self.validate_inputs()
        if headers is None or csv_file_path is None:
            return

        self.clear_not_updated_aliquots()

        csv_rows = self.read_csv(csv_file_path)

        merger = PdfMerger()
        files_added = 0

        for csv_row in csv_rows:
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
        Hospital_Name = self.get_hospital_name(Sample_Collection_Site, Master_ID)
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
                response = requests.post(
                    aliquot_url,
                    json=aliquot_payload,
                    headers=headers,
                    verify=self.cert_path,
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
                        response = requests.post(
                            f"{self.base_url}/aliquots/{pk_aliquot_uid}",
                            json=IDEXX_aliquot_payload,
                            headers=headers,
                            verify=self.cert_path,
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

            frozen_culture_label_url = f"{self.base_url}/labels/3/print"
            frozen_culture_label_payload = {
                "aliquots": frozen_culture_labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    frozen_culture_label_url,
                    json=frozen_culture_label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                frozen_culture_pdf = BytesIO(response.content)
                pdf_merger.append(frozen_culture_pdf)
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label making for SL0 Number {Master_ID}: {e}",
                )
                return

            IDEXX_label_url = f"{self.base_url}/labels/7/print"
            IDEXX_label_payload = {
                "aliquots": IDEXX_labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    IDEXX_label_url,
                    json=IDEXX_label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                IDEXX_pdf = BytesIO(response.content)
                pdf_merger.append(IDEXX_pdf)
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label making for SL0 Number {Master_ID}: {e}",
                )
                return

            merged_label_pdf = BytesIO()
            pdf_merger.write(merged_label_pdf)
            pdf_merger.close()
            merged_label_pdf.seek(0)
            self.log(f"Labels made for SL0 Number {Master_ID}")
            return merged_label_pdf.getvalue()

        else:
            frozen_culture_label_url = f"{self.base_url}/labels/3/print"
            frozen_culture_label_payload = {
                "aliquots": frozen_culture_labels_to_print_ids,
                "numberOfLabelsPerAliquot": 1,
            }
            try:
                response = requests.post(
                    frozen_culture_label_url,
                    json=frozen_culture_label_payload,
                    headers=headers,
                    verify=self.cert_path,
                )
                response.raise_for_status()
                self.log(f"Labels made for SL0 Number {Master_ID}")
            except requests.exceptions.RequestException as e:
                self.log(
                    f"Error during label making for SL0 Number {Master_ID}: {e}",
                )
            return response.content

    def update_aliquots(self):
        headers, csv_file_path = self.validate_inputs()
        if headers is None or csv_file_path is None:
            return

        self.clear_not_updated_aliquots()

        csv_rows = self.read_csv(csv_file_path)

        for csv_row in csv_rows:
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
            output_str = requests.get(url, headers=headers, verify=self.cert_path)
            output = output_str.json()

            position1_value = output["properties"].get("position1")

            if not position1_value or position1_value == "null":
                if (not shelf and not row) and all(
                    s.isdigit() for s in [rack, box, position]
                ):
                    payload = {
                        "fwLocation": {
                            "FK_FreezerSectID": 10016,
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
                response = requests.post(
                    url, json=payload, headers=headers, verify=self.cert_path
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
        tag = "bold" if bold else None
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)

    def open_url(self, url):
        webbrowser.open_new(url)


if __name__ == "__main__":
    root = tk.Tk()
    app = AliquotUpdaterApp(root)
    root.mainloop()
