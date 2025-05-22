# Bulk Continuing Education Certificate Generator (from CSV/Excel to DOCX & PDF)

This Python script automates the generation of personalized DOCX documents using data from a CSV or Excel file, and then converts these documents to PDF format. It uses Tkinter to provide a simple Graphical User Interface (GUI) for file and directory selection.

## Features

* Reads recipient data from a CSV or Excel file.
* Preprocesses data by converting column headers to lowercase and replacing spaces with underscores.
* Adds additional columns to the data, such as an ACE provider name, provider number, and the current date.
* Formats the 'event_date' column to MM-DD-YYYY.
* For each row in the input file, generates a DOCX file using a Word template named `ce_template.docx`.
* Saves the generated DOCX files to a `./output/` directory with filenames formatted as `[name]-[event_date].docx`.
* Batch converts all DOCX files in a user-specified output folder to PDF.

## Prerequisites

* Python 3.x
* The following Python libraries:
    * `pandas`
    * `python-docx-template`
    * `tkinter` (usually comes with Python)
    * `docx2pdf`

## Installation

1.  **Clone the repository or download the script:**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```
    Otherwise, just save the Python script (`your_script_name.py`) to a directory on your computer.

2.  **Install required libraries:**
    Open your terminal or command prompt and run:
    ```bash
    pip install pandas python-docx-template docx2pdf
    ```

3.  **Prepare the template file:**
    * Create a Microsoft Word document named `ce_template.docx`. This file should be in the same directory as the Python script.
    * In this template, use Jinja2-style placeholders (e.g., `{{ column_name }}`) to indicate where data should be inserted. For example:
        * `{{ name }}`
        * `{{ event_date }}`
        * `{{ any_other_column_from_your_csv_excel }}`
        * `{{ ace_name }}` (added by the script)
        * `{{ provider_number }}` (added by the script)
        * `{{ date }}` (added by the script, current date in MM-DD-YYYY format)

4.  **Prepare your input data file:**
    * Your input file should be in CSV or Excel (.xlsx) format.
    * Ensure it contains at least a column named `name` (used for naming output files) and a column named `event_date`. The script will convert `event_date` to `MM-DD-YYYY` format.
    * Other columns can be included based on the placeholders in your `ce_template.docx`.
    * Column headers will be converted to lowercase, and spaces will be replaced with underscores (e.g., "Event Name" becomes "event_name"). Adjust placeholders in your template accordingly.

## How to Use

1.  **Run the script:**
    Navigate to the directory where the script is located using your terminal or command prompt and run:
    ```bash
    python your_script_name.py
    ```
    (Replace `your_script_name.py` with the actual name of your Python file.)

2.  **Select Input File:**
    * A window titled "File and Directory Selection" will appear.
    * Click the "Select File" button.
    * A file dialog will open. Navigate to your CSV or Excel data file, select it, and click "Open". The path to the selected file will be printed to the console.

3.  **Select Save Directory:**
    * Click the "Select Save Directory" button.
    * A directory selection dialog will open. Navigate to the folder where you want the final PDF files to be saved, select it, and click "Select Folder". The path to the selected directory will be printed to the console.
    * After selecting the directory, the GUI window will close automatically.

4.  **Processing:**
    * The script will read the input file.
    * It will process the data, generate a DOCX file for each row, and save it in a subfolder named `output` within the script's directory. If the `output` folder doesn't exist, the script might not work as expected unless it's created or the path is modified.
        * **Note:** The script currently saves DOCX files to `./output/`. However, the PDF conversion step (`convert(out_folder_path)`) converts DOCX files from the `out_folder_path` selected by the user via the GUI. For consistent behavior, you might want to save DOCX files to `out_folder_path` as well.
    * After processing all records, it will convert all `.docx` files in the specified `out_folder_path` to PDF.

5.  **Output:**
    * Intermediate DOCX files will be located in the `./output/` folder (relative to the script's location). Filenames will follow the pattern `[name]-[event_date].docx` (e.g., `John_Doe-05-14-2025.docx`).
    * The final PDF files will be in the directory you selected via the GUI.

## Script Breakdown

* **`select_file()`:** Opens a file dialog to choose the input data file (CSV or Excel).
* **`select_save_directory()`:** Opens a directory dialog to choose the output folder for the final PDF files.
* **Tkinter GUI Setup:** Initializes the main window and adds buttons for file and directory selection.
* **Data Reading:**
    * Tries to read the selected file as a CSV.
    * If a `UnicodeDecodeError` occurs, it attempts to read it as an Excel file.
* **Data Preprocessing:**
    * Converts all column headers to lowercase.
    * Right-strips trailing spaces from each header.
    * Replaces any spaces in column headers with underscores (`_`).
* **Adding/Modifying Columns:**
    * Adds a new column `ace_name` with the hardcoded value 'ace provider name'.
    * Adds a new column `provider_number` with the hardcoded value 'provider number'.
    * Converts the `event_date` column to datetime objects (handling errors), then formats it as an `MM-DD-YYYY` string.
    * Adds a new column `date` with the current date (formatted as `MM-DD-YYYY`).
    * Converts all data types in the DataFrame to string to ensure compatibility with the template.
* **Document Generation:**
    * Iterates through each row of the DataFrame (as a dictionary).
    * Loads the `ce_template.docx` template.
    * Renders the template with the row's data.
    * Cleans up the `name` field for the filename by replacing spaces with underscores.
    * Saves the rendered document to the `./output/` directory with the filename `[name]-[event_date].docx`.
* **PDF Conversion:**
    * Uses the `docx2pdf` library to convert all `.docx` files in the previously selected `out_folder_path` to PDF.

## Important Notes & Customization

* **Hardcoded Values:** The values for `ace_name` ('ace provider name') and `provider_number` ('provider number') are hardcoded directly in the script. You will need to modify the script if you require different values or want to source them from the input file or a configuration.
* **Template File Location:** The script expects `ce_template.docx` to be in the same directory as the script.
* **DOCX Output Directory:** DOCX files are hardcoded to be saved to `./output/`. You might need to create this folder if it doesn't exist, or modify the script to use a different path (e.g., the same `out_folder_path` as the PDF output).
* **Error Handling:** The script includes basic error handling for file reading (`UnicodeDecodeError`). More robust error handling (e.g., file not found, missing columns, template issues) could be added depending on your needs.
* **PDF Conversion Path:** Ensure the directory you select in the GUI for saving PDFs (`out_folder_path`) is the one that contains the DOCX files generated by the script prior to the PDF conversion step. If DOCX files are saved in `./output/` and `out_folder_path` is somewhere else, `convert(out_folder_path)` might not find the files to convert unless you manually move or copy the DOCX files there before the script runs the conversion. Consider unifying output paths for simplicity.

## Potential Improvements

* Allow the user to specify `ace_name` and `provider_number` via the GUI or a configuration file.
* Allow the user to specify the template file via the GUI.
* Allow the user to specify the output directory for intermediate DOCX files via the GUI and unify it with the PDF conversion path.
* Provide more detailed error messages for missing columns or files.
* Add an option to delete intermediate DOCX files after PDF generation.
