
# PDF Form Filler

This project automates the process of filling out large number of PDF forms using data from an Excel file, one form per row. It reads data from the Excel spreadsheet and populates the corresponding fields in a PDF form template. The script supports handling of text fields, date fields, and checkboxes, allowing for user-friendly input (e.g., "Yes" or "No") in the Excel file. The filled PDFs are saved in an output directory for easy access.

## Features

- **Excel to PDF Automation**: Automatically fill PDF forms with data from an Excel spreadsheet.
- **Checkbox Handling**: Process checkbox fields in PDFs based on user-friendly inputs like "Yes" or "No".
- **Date Formatting**: Automatically format date fields to a specified format (e.g., MM/DD/YYYY).
- **Logging**: Comprehensive logging for easy debugging and monitoring.
- **Customizable Mappings**: Easily configure field mappings between Excel columns and PDF form fields.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
  - [Using pip and requirements.txt](#using-pip-and-requirementstxt)
  - [Using conda](#using-conda)
- [Usage](#usage)
  - [Configuration](#configuration)
    - [Field Mappings](#field-mappings)
    - [Checkbox Mappings](#checkbox-mappings)
    - [Date Fields](#date-fields)
  - [Running the Script](#running-the-script)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Prerequisites

- **Python**: Version 3.7 or higher is required.
- **Packages**:
  - `pypdf` (version 3.8.1 or higher)
  - `pandas` (version 1.1.0 or higher)

## Installation

### Using pip and requirements.txt

1. **Clone or Download the Repository**

   ```bash
   git clone https://github.com/WillHsiaoNYC/ExcelFillPDF.git
   cd ExcelFillPDF
   ```

2. **Create a Virtual Environment (Recommended)**

   ```bash
   # For Windows
   python -m venv venv
   venv\Scripts\activate

   # For macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Upgrade pip (Optional)**

   ```bash
   pip install --upgrade pip
   ```

4. **Install Required Packages**

   ```bash
   pip install -r requirements.txt
   ```

   **Contents of `requirements.txt`:**

   ```plaintext
   pypdf>=3.8.1
   pandas>=1.1.0
   ```

### Using conda

1. **Create a New Conda Environment**

   ```bash
   conda create -n pdf_env python=3.8
   ```

2. **Activate the Environment**

   ```bash
   conda activate pdf_env
   ```

3. **Install Required Packages**

   ```bash
   conda install pandas
   conda install -c conda-forge pypdf
   ```

## Usage

### Configuration

Before running the script, you need to configure it to match your specific PDF form and Excel file.

#### Field Mappings

Edit the `FIELD_MAPPING` dictionary in the script to map your Excel columns to the PDF text fields.

```python
# Regular field mapping from Excel columns to PDF field names
FIELD_MAPPING = {
    'Name': 'pdf_name_field',
    'Date of Birth': 'pdf_dob_field',
    'Email': 'pdf_email_field',
    # Add other mappings here
}
```

#### Checkbox Mappings

Edit the `CHECKBOX_MAPPINGS` dictionary to handle checkbox fields. The script allows users to enter "Yes" or "No" in the Excel file, which are then mapped to the corresponding checkboxes in the PDF.

```python
# Checkbox mappings
CHECKBOX_MAPPINGS = {
    'Subscribe to Newsletter': {
        'yes': {
            'fields': {
                'newsletter_yes_checkbox': '/Yes',
                'newsletter_no_checkbox': '',
            }
        },
        'no': {
            'fields': {
                'newsletter_yes_checkbox': '',
                'newsletter_no_checkbox': '/No',
            }
        }
    },
    # Add other checkbox mappings here
}
```

- **Excel Column**: `'Subscribe to Newsletter'`
- **Possible Values**: `'Yes'` or `'No'`
- **PDF Checkbox Fields**: `'newsletter_yes_checkbox'` and `'newsletter_no_checkbox'`

#### Date Fields

Specify any date fields that need special formatting in the `DATE_FIELDS` list.

```python
# List of Excel columns that are date fields
DATE_FIELDS = ['Date of Birth', 'Hire Date']
```

### Running the Script

1. **Place Your Excel File and PDF Template in the Project Directory**

   - **Excel File**: For example, `data.xlsx`
   - **PDF Template**: For example, `form_template.pdf`

2. **Update File Paths in the Script**

   In the `main` function, update the file paths to point to your Excel file and PDF template.

   ```python
   def main():
       # File paths
       excel_file = Path("data.xlsx")  # Replace with your Excel file path
       template_pdf = Path("form_template.pdf")  # Replace with your PDF form path
       output_folder = Path("output")  # Replace with your desired output folder path
       # ...
   ```

3. **Run the Script**

   ```bash
   python ExcelFillPDF.py
   ```

4. **Output**

   - The filled PDFs will be saved in the specified output directory (default is `output`).
   - Each PDF file will be named using the `'filename'` column from the Excel file, if available; otherwise, it will use a default naming convention like `output_1.pdf`, `output_2.pdf`, etc.

## Logging

- The script uses the `logging` module to output messages.
- Logging levels can be adjusted in the script configuration:

  ```python
  # Configure logging
  logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
  ```

- **Logging Levels**:
  - `DEBUG`: Detailed information, typically of interest only when diagnosing problems.
  - `INFO`: Confirmation that things are working as expected.
  - `WARNING`: An indication that something unexpected happened.
  - `ERROR`: Due to a more serious problem, the software has not been able to perform some function.
  - `CRITICAL`: A serious error, indicating that the program itself may be unable to continue running.

## Troubleshooting

- **Field Names Not Matching**:

  - Use the `inspect_pdf_fields` function to print out all the field names in your PDF form.
  - Uncomment the `inspect_pdf_fields(template_pdf)` line in the `main` function.

    ```python
    def main():
        # ...
        # Uncomment the next line to inspect PDF fields
        inspect_pdf_fields(template_pdf)
        # ...
    ```

- **Checkbox Values Not Setting Correctly**:

  - Ensure that the checkbox values in `CHECKBOX_MAPPINGS` match the export values in the PDF form.
  - Use the `inspect_pdf_fields` function to determine the correct values.

- **Excel Columns Not Found**:

  - Verify that the column names in your Excel file match exactly with those specified in `FIELD_MAPPING` and `CHECKBOX_MAPPINGS`.
  - Check for leading/trailing spaces or case differences.

- **Date Parsing Errors**:

  - Ensure that date fields in your Excel file are formatted correctly or can be parsed by `pandas.to_datetime`.
  - If necessary, pre-process your Excel data to standardize date formats.

- **Python Version Issues**:

  - Confirm that you are using Python 3.7 or higher.
  - Check your Python version:

    ```bash
    python --version
    ```

- **Module Not Found Errors**:

  - Make sure all required packages are installed in your environment.
  - If using a virtual environment, ensure it is activated before running the script.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- **[pypdf](https://pypdf.readthedocs.io/en/latest/)**: A pure Python PDF library capable of splitting, merging, cropping, and transforming PDF files.
- **[pandas](https://pandas.pydata.org/)**: An open-source data analysis and manipulation tool.

