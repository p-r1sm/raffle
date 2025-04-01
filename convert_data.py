import pandas as pd
import os
import sys
import math


def capitalize_words(text):
    """
    Capitalize each word in a string, handling None values and non-string types.
    
    Args:
        text: Input text to capitalize
    Returns:
        Capitalized string with each word capitalized
    """
    if pd.isna(text):  # Handle None/NaN values
        return ""
    return " ".join(word.capitalize() for word in str(text).split())


def process_csv_in_batches(input_file, output_prefix="processed_data", batch_size=300):
    """
    Process CSV/Excel data with specific column transformations and save in batches.

    Args:
        input_file (str): Path to the input file (CSV or Excel)
        output_prefix (str): Prefix for output CSV files
        batch_size (int): Number of rows per batch (including header)
    """
    try:
        # Check file extension
        file_extension = os.path.splitext(input_file)[1].lower()
        
        # Read the input file based on its extension
        if file_extension in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(input_file)
                print(f"Successfully read Excel file")
            except Exception as e:
                raise ValueError(f"Error reading Excel file: {str(e)}")
        else:
            # Try different encodings for CSV
            encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(input_file, encoding=encoding)
                    print(f"Successfully read CSV file using {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise ValueError("Could not read the CSV file with any of the attempted encodings")

        # Verify required columns exist
        required_columns = {
            "id": ["id", "ID", "Id", "serial", "Serial No", "serial_no"],
            "first_name": ["first_name", "firstname", "fname", "First Name"],
            "last_name": ["last_name", "lastname", "lname", "Last Name"],
            "arpit_group": ["arpit_group", "arpitgroup", "group", "Arpit group"],
            "area": ["area", "location", "place", "Area"],
            "intl_code": [
                "Int'l Calling code (e.g. US 1, UK 44)",
            ],
            "whatsapp": ["whatsapp", "whatsapp_number", "WhatsApp Number"],
        }

        # Find matching columns for each required field
        column_mapping = {}
        for required_col, possible_names in required_columns.items():
            found = False
            for name in possible_names:
                matching_cols = [
                    col for col in df.columns if name.lower() == col.lower()
                ]
                if matching_cols:
                    column_mapping[required_col] = matching_cols[0]
                    found = True
                    break

            if not found:
                print(f"Warning: Could not find column for {required_col}")
                print(f"Available columns: {list(df.columns)}")
                user_input = input(
                    f"Enter the column name to use for {required_col}: "
                ).strip()
                if user_input in df.columns:
                    column_mapping[required_col] = user_input
                else:
                    raise ValueError(f"Invalid column name provided for {required_col}")

        # Create transformed columns
        processed_df = pd.DataFrame()

        # Add ID column first
        processed_df["ID"] = df[column_mapping["id"]].astype(str)

        # Combine and capitalize names, handling multiple words
        first_names = df[column_mapping["first_name"]].apply(capitalize_words)
        last_names = df[column_mapping["last_name"]].apply(capitalize_words)
        processed_df["NAME"] = first_names + " " + last_names

        # Capitalize area (handle multiple words)
        processed_df["AREA"] = df[column_mapping["area"]].apply(capitalize_words)

        # Convert Arpit Group to uppercase
        processed_df["ARPIT_GROUP"] = df[column_mapping["arpit_group"]].str.strip().str.upper()

        # Combine international code and whatsapp number
        # First, ensure the numbers are strings and remove any non-numeric characters
        intl_codes = (
            df[column_mapping["intl_code"]]
            .astype(str)
            .str.extract('(\d+)', expand=False)
            .fillna('')
        )
        whatsapp_numbers = (
            df[column_mapping["whatsapp"]]
            .astype(str)
            .str.extract('(\d+)', expand=False)
            .fillna('')
        )
        processed_df["CONTACT_NUMBER"] = "+" + intl_codes + whatsapp_numbers

        # Calculate number of batches needed
        total_rows = len(processed_df)
        num_batches = math.ceil(
            total_rows / (batch_size - 1)
        )  # -1 to account for header

        # Save in batches
        for i in range(num_batches):
            start_idx = i * (batch_size - 1)
            end_idx = min((i + 1) * (batch_size - 1), total_rows)

            # Create batch filename
            batch_file = f"{output_prefix}_batch_{i + 1}.csv"

            # Save batch
            batch_df = processed_df.iloc[start_idx:end_idx]
            batch_df.to_csv(batch_file, index=False)
            print(f"Saved batch {i + 1} to {batch_file} ({len(batch_df)} records)")

        print(f"Processing complete. Created {num_batches} batch files.")
        return True

    except Exception as e:
        print(f"Error processing data: {str(e)}")
        return False


if __name__ == "__main__":
    # Check if the script is called with an input file
    if len(sys.argv) < 2:
        print("Usage: python convert_data.py input_file.csv [output_prefix]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_prefix = sys.argv[2] if len(sys.argv) > 2 else "processed_data"

    process_csv_in_batches(input_file, output_prefix)
