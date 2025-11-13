# PDF to Excel Conversion Script (Hierarchical)
#
# Description:
# This script extracts hierarchical, tabular data from a PDF file and saves it
# into a structured Excel spreadsheet. It identifies different levels in the
# document structure and parses data rows based on a specific "X X" delimiter.
#
# Required Libraries:
# - PyMuPDF (fitz): For reading and extracting text from PDF files.
# - pandas: For data manipulation and creating the Excel file.
# - openpyxl: Required by pandas to write .xlsx files.

import fitz  # PyMuPDF
import pandas as pd
import re
import os

def extract_text_from_pdf(pdf_path):
    """
    Extracts all text content from a given PDF file.

    Args:
        pdf_path (str): The file path to the PDF.

    Returns:
        str: A single string containing all the text from the PDF,
             with pages separated by newlines.
             Returns None if the file is not found.
    """
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found.")
        return None

    try:
        doc = fitz.open(pdf_path)
        full_text = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Using 'blocks' can sometimes preserve layout better than get_text()
            blocks = page.get_text("blocks")
            blocks.sort(key=lambda b: (b[1], b[0])) # Sort by top then left coordinate
            for b in blocks:
                full_text.append(b[4].strip()) # b[4] is the text content
        return "\n".join(full_text)
    except Exception as e:
        print(f"An error occurred while reading the PDF: {e}")
        return None

def parse_text_to_dataframe(text):
    """
    Parses the extracted text with hierarchical levels into a pandas DataFrame.
    It identifies headers for different levels and associates data rows with them.

    Args:
        text (str): The raw text extracted from the PDF.

    Returns:
        pd.DataFrame: A DataFrame containing the parsed and structured data.
    """
    # Regex patterns to identify different line types
    # Level 1 Header: e.g., "001 Net Revenue"
    level1_pattern = re.compile(r'^\d{3}\s+.*')
    # Level 2 Header: e.g., "001.1 Service Revenue"
    level2_pattern = re.compile(r'^\d{3}\.\d+\s+.*')
    # Level 3 Header: e.g., "4010 Gross Service Revenue"
    level3_pattern = re.compile(r'^\d{4}\s+.*')
    # Data Row: Contains CADT code, numeric code, and "X X" delimiter
    # It captures the CADT code (optional), numeric code (optional), and the description after "X X"
    data_pattern = re.compile(r'^(CADT\w+)?\s*-\s*(\d{8,})?\s*_?\s*X\s*_?\s*(.*)')

    data = []
    lines = text.split('\n')

    # State variables to hold the current hierarchy context
    current_level_1 = ""
    current_level_2 = ""
    current_level_3_code = ""
    current_level_3_desc = ""

    for line in lines:
        line = line.strip()
        if not line or "PAGE" in line or "Deloitte India Management Reporting" in line:
            continue

        # Check for headers first to update the state
        if level1_pattern.match(line) and not level2_pattern.match(line) and not level3_pattern.match(line):
            current_level_1 = line.split(maxsplit=1)[1] if len(line.split(maxsplit=1)) > 1 else line
            current_level_2 = "" # Reset sublevels
            current_level_3_code = ""
            current_level_3_desc = ""
        elif level2_pattern.match(line):
            current_level_2 = line.split(maxsplit=1)[1] if len(line.split(maxsplit=1)) > 1 else line
            current_level_3_code = "" # Reset sublevel
            current_level_3_desc = ""
        elif level3_pattern.match(line):
            parts = line.split(maxsplit=1)
            current_level_3_code = parts[0]
            current_level_3_desc = parts[1] if len(parts) > 1 else ""
        else:
            # If not a header, try to match as a data row
            match = data_pattern.match(line)
            if match:
                id_code, numeric_code, description = match.groups()

                # Clean up the extracted data
                id_code = id_code if id_code else ''
                numeric_code = numeric_code if numeric_code else ''
                description = description.strip()

                # Heuristic to find numeric code if it was missed by the regex
                if id_code and not numeric_code:
                    desc_parts = description.split(maxsplit=1)
                    if len(desc_parts) > 1 and desc_parts[0].isdigit() and len(desc_parts[0]) >= 8:
                        numeric_code = desc_parts[0]
                        description = desc_parts[1]

                # Append the structured data row
                if id_code or numeric_code: # Ensure it's a valid data row
                    data.append([
                        current_level_1,
                        current_level_2,
                        current_level_3_code,
                        current_level_3_desc,
                        id_code,
                        numeric_code,
                        description
                    ])

    # Create a DataFrame with columns matching the user's format
    columns = [
        'Level 1', 'Level 2', 'Level 3 Code', 'Level 3',
        'ID', 'Numeric_Code', 'Description'
    ]
    df = pd.DataFrame(data, columns=columns)
    return df

def save_to_excel(dataframe, excel_path):
    """
    Saves a pandas DataFrame to an Excel file.

    Args:
        dataframe (pd.DataFrame): The DataFrame to save.
        excel_path (str): The desired output path for the Excel file.
    """
    try:
        dataframe.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"Successfully converted PDF to Excel file: '{excel_path}'")
    except Exception as e:
        print(f"An error occurred while saving to Excel: {e}")

def convert_pdf_to_excel(pdf_path, excel_path):
    """
    Main function to convert a PDF file to Excel.

    Args:
        pdf_path (str): Path to the input PDF file.
        excel_path (str): Path to the output Excel file.

    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        # Step 1: Extract text from the PDF
        raw_text = extract_text_from_pdf(pdf_path)

        if not raw_text:
            return False, "Failed to extract text from the PDF."

        # Step 2: Parse the text into a structured DataFrame
        data_df = parse_text_to_dataframe(raw_text)

        if data_df.empty:
            return False, "Could not parse any structured data from the PDF."

        # Step 3: Save the DataFrame to an Excel file
        save_to_excel(data_df, excel_path)
        return True, f"Successfully converted PDF to Excel file: '{excel_path}'"

    except Exception as e:
        return False, f"An error occurred during conversion: {str(e)}"

# --- Main Execution ---
if __name__ == "__main__":
    # Define the input PDF and output Excel file names.
    # The script assumes 'FSV India ALL.pdf' is in the same directory.
    pdf_filename = "FSV_MV.pdf"
    excel_filename = "FSV_MV.xlsx"

    print(f"Starting conversion of '{pdf_filename}'...")

    # Step 1: Extract text from the PDF
    raw_text = extract_text_from_pdf(pdf_filename)

    if raw_text:
        # Step 2: Parse the text into a structured DataFrame
        print("Parsing extracted text with hierarchical logic...")
        data_df = parse_text_to_dataframe(raw_text)

        if not data_df.empty:
            # Step 3: Save the DataFrame to an Excel file
            print(f"Saving data to '{excel_filename}'...")
            save_to_excel(data_df, excel_filename)
        else:
            print("Could not parse any structured data from the PDF.")
    else:
        print("Failed to extract text from the PDF. Aborting conversion.")