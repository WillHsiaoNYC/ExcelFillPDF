import pandas as pd
import logging
from pathlib import Path
from typing import Optional
import pypdf
from pypdf.generic import NameObject, BooleanObject, IndirectObject

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Regular field mapping from Excel columns to PDF field names
FIELD_MAPPING = {
    'Name(s)': 'f1_01[0]',  # Example mapping
    'Identifying number': 'f1_02[0]',
    'Year': 'f1_03[0]',
    'Make': 'f1_04[0]',
    'Model': 'f1_05[0]',
    'VIN': 'f1_06[0]',
    'Date in Service': 'f1_07[0]',
    # Exclude checkboxes from FIELD_MAPPING
    # Add other mappings here
}

# Checkbox mappings
CHECKBOX_MAPPINGS = {
    'more than $25,000?': {
        'yes': {
            'fields': {
                'c1_1[0]': '/1',  # Yes checkbox checked
                'c1_1[1]': '',    # No checkbox unchecked
            }
        },
        'no': {
            'fields': {
                'c1_1[0]': '',    # Yes checkbox unchecked
                'c1_1[1]': '/2',  # No checkbox checked
            }
        }
    },
    'new clean vehicle?': {
        'yes': {
            'fields': {
                'c1_2[0]': '/1',  # Yes checkbox checked
                'c1_2[1]': '',    # No checkbox unchecked
            }
        },
        'no': {
            'fields': {
                'c1_2[0]': '',    # Yes checkbox unchecked
                'c1_2[1]': '/2',  # No checkbox checked
            }
        }
    },
    # Add other checkbox mappings here
}

# List of Excel columns that are date fields
DATE_FIELDS = ['Date in Service']  # Add other date fields as needed


def read_excel_file(file_path: Path) -> pd.DataFrame:
    """
    Reads an Excel file and returns a DataFrame.
    :param file_path: Path to the Excel file
    :return: pandas DataFrame
    """
    try:
        df = pd.read_excel(file_path)
        logging.info("Excel file loaded successfully!")
        return df
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        raise


def inspect_pdf_fields(pdf_path: Path) -> None:
    """
    Inspects the fields of a PDF form and logs the field names, details, and current values.
    :param pdf_path: Path to the PDF form
    """
    try:
        pdf_reader = pypdf.PdfReader(str(pdf_path))
        form_fields = pdf_reader.get_fields()

        if form_fields:
            for field_name, field_info in form_fields.items():
                # Log the field name
                logging.info(f"Field Name: {field_name}")

                # Log the field details (metadata)
                logging.info(f"Field Info: {field_info}")

                # Get and log the current value of the field, if any
                current_value = field_info.get('/V', 'No value set')
                logging.info(f"Field Value: {current_value}")

                logging.info("-" * 40)
        else:
            logging.info("No form fields found in the PDF.")
    except Exception as e:
        logging.error(f"Error reading PDF: {e}")
        raise


def ensure_output_folder(output_folder: Path) -> None:
    output_folder.mkdir(parents=True, exist_ok=True)


def set_need_appearances_writer(writer: pypdf.PdfWriter):
    # Set the /NeedAppearances flag in the AcroForm dictionary
    try:
        catalog = writer._root_object
        if "/AcroForm" not in catalog:
            # Create an AcroForm if one doesn't exist
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)
            })
        acro_form = catalog["/AcroForm"]
        acro_form.update({NameObject("/NeedAppearances"): BooleanObject(True)})
    except Exception as e:
        logging.error(f"Error setting NeedAppearances: {e}")
        raise


def fill_pdf_with_data(
    template_pdf_path: Path,
    output_folder: Path,
    df: pd.DataFrame,
    field_mapping: dict,
    checkbox_mappings: dict,
    date_fields: list
) -> None:
    """
    Fills a PDF form using data from an Excel DataFrame and saves each PDF.
    :param template_pdf_path: Path to the blank PDF form
    :param output_folder: Folder where the filled PDFs will be saved
    :param df: DataFrame containing the data from Excel
    :param field_mapping: Mapping of Excel columns to PDF field names
    :param checkbox_mappings: Mapping of Excel columns to checkbox options
    :param date_fields: List of Excel columns that are date fields
    """
    ensure_output_folder(output_folder)

    try:
        for index, row in df.iterrows():
            reader = pypdf.PdfReader(str(template_pdf_path))
            writer = pypdf.PdfWriter()

            # Copy pages from reader to writer
            for page in reader.pages:
                writer.add_page(page)

            # Copy the AcroForm from reader to writer and set NeedAppearances
            if "/AcroForm" in reader.trailer["/Root"]:
                writer._root_object.update({
                    NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
                })
                set_need_appearances_writer(writer)
            else:
                logging.warning("No AcroForm found in the PDF.")

            # Prepare a dictionary to hold the field values
            field_values = {}

            # Process regular fields
            for excel_column, pdf_field_name in field_mapping.items():
                if excel_column in df.columns:
                    field_value = row[excel_column]
                    if pd.isnull(field_value):
                        field_value = ''
                    else:
                        if excel_column in date_fields:
                            # Format the date to MM/DD/YYYY
                            try:
                                field_value = pd.to_datetime(field_value).strftime('%m/%d/%Y')
                            except Exception as e:
                                logging.warning(f"Date parsing failed for '{field_value}': {e}")
                                field_value = str(field_value)
                        else:
                            field_value = str(field_value)
                    field_values[pdf_field_name] = field_value
                else:
                    logging.warning(f"Field '{excel_column}' not found in DataFrame columns")

            # Process checkbox fields
            for excel_column, options in checkbox_mappings.items():
                if excel_column in df.columns:
                    field_value = row[excel_column]
                    field_value_str = str(field_value).strip().lower()
                    checkbox_info = options.get(field_value_str)
                    if checkbox_info and 'fields' in checkbox_info:
                        fields_to_set = checkbox_info['fields']
                        field_values.update(fields_to_set)
                    else:
                        logging.warning(f"Unrecognized checkbox value '{field_value}' for field '{excel_column}'")
                else:
                    logging.warning(f"Checkbox field '{excel_column}' not found in DataFrame columns")

            # Update the field values in the PDF
            writer.update_page_form_field_values(writer.pages[0], field_values)

            # Save the filled PDF using the filename from a column in the DataFrame, e.g., 'filename'
            if 'filename' in df.columns:
                output_filename = output_folder / f"{row['filename']}.pdf"
            else:
                output_filename = output_folder / f"output_{index + 1}.pdf"

            with output_filename.open('wb') as output_pdf:
                writer.write(output_pdf)
            logging.info(f"PDF saved as {output_filename}")

    except Exception as e:
        logging.error(f"Error filling PDF: {e}")
        raise


def main():
    # File paths
    excel_file = Path("PDF Source.xlsx")  # Replace with your Excel file path
    template_pdf = Path("Form 8936 Schedule A Clean Vehicle Credit Amount (2023).pdf")  # Replace with your PDF form path
    output_folder = Path("output")  # Replace with your desired output folder path

    # Step 1: Read the Excel file
    try:
        df = read_excel_file(excel_file)
        logging.info(f"DataFrame Head:\n{df.head()}")
    except Exception:
        return

    # Step 2: Inspect the PDF form fields (optional)
    # Uncomment the next line to inspect PDF fields
    # inspect_pdf_fields(template_pdf)

    # Step 3: Fill the PDF with data from Excel
    logging.info("Filling PDF forms with data from Excel...")
    try:
        fill_pdf_with_data(template_pdf, output_folder, df, FIELD_MAPPING, CHECKBOX_MAPPINGS, DATE_FIELDS)
    except Exception:
        return


if __name__ == "__main__":
    main()
