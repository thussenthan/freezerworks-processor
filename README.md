# Freezerworks Processor

[![GitHub](https://img.shields.io/badge/GitHub-FreezerworksProcessor-blue?style=flat&logo=github)](https://github.com/yourusername/Freezerworks-Processor)

**Version 1.5** (Updated as of 02/04/2025)

## Overview

Freezerworks Processor is a Python-based GUI application designed to automate various processes within the Freezerworks system. The application provides three primary functionalities:
- **Process Patient Sample** – Processes sample data and prints PDF labels.
- **Freeze Passaged Cells** – Manages aliquot creation for passaged cells with associated PDF label generation.
- **Aliquot Freezer Assignment** – Updates aliquot freezer location assignments using CSV input.

The program uses Bearer Token authentication to securely connect with the Freezerworks REST API, and it includes features such as CSV template generation, SSL certificate verification, and real-time logging.

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
- **SSL Certificate Support:** Ensures secure API connections using a provided certificate.
- **Logging:** Displays real-time messages, errors, and progress updates within the GUI.
- **PDF Merging:** Uses PyPDF2 to combine generated PDF labels.

## Prerequisites

- **Python 3.x**  
- **Required Python Packages:**  
  Install using pip:
  ```bash
  pip install requests PyPDF2
  ```
- **Tkinter:** Usually bundled with Python.
- **SSL Certificate:**  
  Place the file `freezerworks.pennstatehealth.net.crt` or something similar in the same directory as the script/executable.
- **Access to the Freezerworks API:**  
  Ensure you have valid API credentials and network access.

## Setup

1. **Clone or Download the Repository:**
   ```bash
   git clone https://github.com/thussenthan/Freezerworks-Processor.git
   ```
2. **Verify the SSL Certificate:**  
   Ensure that `freezerworks.pennstatehealth.net.crt` (or any SSL certificate, if applicable) is in the repository directory.
3. **Run the Application:**
   - **Executable:**
     To make an executable file, run the following command:
     ```bash
     pyinstaller --clean --onefile --windowed --add-data "freezerworks.pennstatehealth.net.crt;." "freezerworks_processor.py"
     ```
     Then double-click `Freezerworks Processor.exe` to launch.
   - **Python Script:**  
     Run the following command:
     ```bash
     python FreezerworksProcessor.py
     ```

## Usage

1. **Obtain Your Bearer Token:**
   - Click the blue **"Enter your Bearer Token"** hyperlink in the application.
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

6. **Execute the Update/Processing:**
   - Click the **"Update"** button.
   - The application will process your data, update aliquot records via the Freezerworks API, and generate PDF labels where applicable.
   - Monitor the log window for progress messages and any errors.
   - If PDF labels are generated, you will be prompted to save the merged PDF file.

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

## Support & Contributions

If you have any questions, encounter issues, or would like to contribute improvements or bug fixes, please open an issue or submit a pull request on [GitHub](https://github.com/yourusername/Freezerworks-Processor).

## Disclaimer

This software interacts directly with your Freezerworks database via API calls. **Always ensure you have a backup of your data** before running any automated processes. Use this tool at your own risk—the developers are not liable for any data loss or system issues.

---

**Note:** This project is distributed without personal or identifying information and is intended for public use and contribution.

---

Feel free to modify any sections (such as repository URLs, allowed CSV values, or instructions) to best match your project’s needs.
