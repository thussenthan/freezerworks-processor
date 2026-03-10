# Freezerworks Processor

[![GitHub](https://img.shields.io/badge/GitHub-FreezerworksProcessor-blue?style=flat&logo=github)](https://github.com/thussenthan/Freezerworks-Processor)

**Version 1.7** (Updated as of 03/10/2026)

## Overview

Freezerworks Processor is a Python GUI that automates common Freezerworks workflows. It supports three functions:
- **Process Patient Sample**: Create aliquots from patient sample metadata and print PDF labels.
- **Freeze Passaged Cells**: Create aliquots for passaged cell samples and print PDF labels.
- **Aliquot Freezer Assignment**: Update aliquot freezer locations from a CSV.

Authentication is via **Bearer Token** against the Freezerworks REST API. The app also supports CSV template generation, SSL certificate verification, and real-time logging.

---

> **Note:** This project was originally developed for lab-specific operations at the Sholler Lab, Penn State College of Medicine, and may require code customization to fit your specific requirements.

---

## Features

- **User-Friendly GUI:** Built with Tkinter for ease of use.
- **Bearer Token Authentication:** Securely authenticate against the Freezerworks API.
- **Multi-Functional:**  
  - *Process Patient Sample* – Search for sample records, create aliquots, and print labels.
  - *Freeze Passaged Cells* – Process cell culture samples and generate multiple PDF labels.
  - *Aliquot Freezer Assignment* – Update the location of aliquots in Freezerworks based on CSV input.
- **CSV Template Generation:** Download preformatted CSV templates for each functionality.
- **SSL Certificate Support:** Uses a provided `.crt` or `.cer` file for TLS verification.
- **Logging:** Displays real-time messages, errors, and progress updates within the GUI.
- **PDF Merging:** Uses PyPDF2 to combine generated PDF labels.

## Quick Start (New Users)

1. **Get a Bearer Token:** In the app, click **"Bearer Token*:"** and log in to Freezerworks.
2. **Pick a workflow:** Choose one of the three radio options.
3. **Download a CSV template:** Click **"Download CSV"** for your workflow.
4. **Delete the example row:** Replace it with your real data.
5. **Select your CSV file:** Click the CSV field and choose your file.
6. **Click Update:** Watch the log window for progress and errors.

## First Run Checklist

- You can open the Freezerworks API docs page from the token link in the app.
- The certificate file `freezerworks.pennstatehealth.net.cer` (or `.crt`) is in the same folder as the script/executable.
- You can open the downloaded CSV template and save it after replacing the example row.
- You have valid Freezerworks credentials with access to the needed tables.

## Prerequisites

- **Python 3.x**  
- **Required Python Packages:**
  ```bash
  python3 -m pip install requests PyPDF2
  ```
- **Tkinter:** Usually bundled with Python on Windows and macOS. If your Python lacks Tk, install the Tk package for your Python distro.
- **SSL Certificate:** The repo includes `freezerworks.pennstatehealth.net.cer`. The app looks for `.crt` or `.cer` in the same directory as the script/executable.
- **Freezerworks API Access:** You need valid credentials and network access.

## Setup

### 1) Clone the Repository

```bash
git clone https://github.com/thussenthan/Freezerworks-Processor.git
cd Freezerworks-Processor
```

### 2) Install Dependencies

```bash
python3 -m pip install --upgrade pip
python3 -m pip install requests PyPDF2
```

### 3) Verify the Certificate

Make sure `freezerworks.pennstatehealth.net.cer` (or `.crt`) is next to the script or executable.

### 4) Run the App (macOS or Windows)

```bash
python3 freezerworks_processor.py
```

## Build Executables

### macOS

```bash
python3 -m pip install pyinstaller
python3 build.py --target macos
```

Launch `dist/freezerworks_processor`.

### Windows

```bash
python -m pip install pyinstaller
python build.py --target windows
```

Launch `Freezerworks Processor.exe` from `dist/`.

## Usage

1. **Obtain Your Bearer Token:**
   - Click the blue **"Bearer Token*:"** label in the application.
   - A web browser will open to the Freezerworks API documentation/login page.
   - Enter your Freezerworks credentials and click **"Get Token"**.
   - Use the clipboard button (cyan/light blue) to copy the full token.
   - Paste the Bearer Token into the token field in the application.

2. **Select the Desired Functionality:**
   - Choose one of the three options:
     - **Process Patient Sample**
     - **Freeze Passaged Cells**
     - **Aliquot Freezer Assignment**

3. **Download the CSV Template:**
   - Click the **"Download CSV"** button next to your chosen task.
   - The template will open automatically (e.g., in Excel).  
   **Important:** Delete or replace the first example row before entering your own data.

4. **Prepare Your CSV File:**
   - Fill in the CSV file with your data according to the template format (see details below).
   - Save the CSV file.

5. **Select Your CSV File:**
   - Click on the file path field (or the white box labeled **"Select CSV File:"**) to browse and choose your saved CSV.
   - If you selected a workflow and no CSV was set, the app will open the file picker automatically.

6. **Execute the Update/Processing:**
   - Click the **"Update"** button.
   - The application will process your data, update aliquot records via the Freezerworks API, and generate PDF labels where applicable.
   - Monitor the log window for progress messages and any errors.
   - If PDF labels are generated, you will be prompted to save the merged PDF file.
   - If the app appears to do nothing after clicking **Update**, check the log file:
     - **Script run:** `freezerworks_processor.log` in the repo directory.
     - **macOS app bundle:** `~/Library/Logs/Freezerworks Processor/freezerworks_processor.log`.

## No Python?

If the user does not have Python or the required packages installed, use the packaged executable:
- **Windows:** Distribute `Freezerworks Processor.exe` from `dist/` (no Python required).
- **macOS:** Distribute `dist/freezerworks_processor` (no Python required).

## Windows "failed to remove temporary directory"

If you see `failed to remove temporary directory` when closing the `.exe`, rebuild with `--runtime-tmpdir` so PyInstaller unpacks to a dedicated folder under `LocalAppData` instead of the default Windows temp location:

```bash
python build.py --target windows
```

## CSV Template Details

### Process Patient Sample CSV

- **Filename:** `Process Patient Sample.csv`  
- **Columns (in order):**
  - *Sample Collection Site* (integer, no leading zeros)
  - *Sample Study ID*
  - *SL0 Number* (four-digit number; do not include "SL0")
  - *Aliquot Type* (must be one of: "Biomarker", "ctDNA", "NK", "PK", or "BMA")
  - *Date of Collection* (MM/DD/YYYY)
  - *Freezing Date* (MM/DD/YYYY)
  - *(Study Time Point)* – Optional
  - *(Notes)* – Optional
  - *(Number of PK Aliquots)* – Must be included if processing PK samples

### Freeze Passaged Cells CSV

- **Filename:** `Freeze Passaged Cells.csv`  
- **Columns (in order):**
  - *SL0 Number* (four-digit number; do not include "SL0")
  - *Cell Line Name*
  - *Sample Study ID*
  - *Aliquot Type*
  - *Date of Collection* (MM/DD/YYYY)
  - *Sample Collection Site* (integer, no leading zeros)
  - *Date of Culture Initiation* (MM/DD/YYYY)
  - *Passage Number* (integer; do not include a "P" prefix)
  - *Freezing Date* (MM/DD/YYYY)
  - *(Media)* – Optional  
    *(Allowed values: "DMEM", "DMEM/F12", "FBS", "MEM Alpha", "Neurocult", "OPTI-MEM", "RPMI", "Tumor Stem Media")*
  - *(Serum Supplement)* – Optional  
    *(Allowed values: "10% FBS", "20% FBS", or "Serum Free")*
  - *(Notes)* – Optional

### Aliquot Freezer Assignment CSV

- **Filename:** `Aliquot Freezer Assignment.csv`  
- **Columns (in order):**
  - *Aliquot ID*
  - *(Shelf)* – Optional
  - *Rack*
  - *(Row)* – Optional
  - *Box*
  - *Position*

> **Note:**  
> - If both **(Shelf)** and **(Row)** are blank, the aliquot is processed for the Liquid Nitrogen Freezer.  
> - If all five fields are provided, the aliquot is processed for the -80 Freezer.

## Freezerworks Processor User Guide

1. **Launch the Application:**  
   Open the program by double-clicking `Freezerworks Processor.exe` (or run the Python script).

2. **Authenticate:**  
   - Click the blue **"Enter your Bearer Token"** link.
   - In the opened web page, enter your Freezerworks credentials and click **"Get Token"**.
   - Use the provided clipboard button to copy the full token.
   - Paste the token into the application’s token field.

3. **Select the Functionality:**  
   Choose one of:
   - Process Patient Sample
   - Freeze Passaged Cells
   - Aliquot Freezer Assignment

4. **Download and Prepare CSV Template:**  
   - Click the corresponding **"Download CSV"** button.
   - Edit the CSV file in your preferred editor (e.g., Excel).  
     **Remember:** Delete or replace the example row before using.

5. **Select Your CSV File:**  
   Click on the **"Select CSV File:"** field to browse and select your saved CSV file.

6. **Execute the Process:**  
   Click **"Update"**.  
   - Watch the log for progress updates and error messages.
   - For processes generating PDF labels, you will be prompted to save the merged PDF file.

7. **Review Output:**  
   Once the process completes, check the log for confirmation messages such as “All aliquots were updated successfully” or lists of any errors encountered.

## Troubleshooting

- **Tkinter not found / GUI won’t launch:** Install the Tk package for your Python distribution.
- **SSL errors (`SSLError`):** Confirm the `.cer`/`.crt` file is next to the script or executable.
- **App opens but nothing happens on Update:** Check the log file path listed above.
- **macOS app blocked by Gatekeeper:** Right-click the app and choose **Open** once.

## Support & Contributions

If you have questions, encounter issues, or would like to contribute improvements, please open an issue or submit a pull request on GitHub.

## Disclaimer

This software interacts directly with your Freezerworks database via API calls. **Always ensure you have a backup of your data** before running any automated processes. Use this tool at your own risk—the developers are not liable for any data loss or system issues.

---

**Note:** This project is distributed without personal or identifying information and is intended for public use and contribution.

---

Feel free to modify any sections (such as repository URLs, allowed CSV values, or instructions) to best match your project’s needs.
