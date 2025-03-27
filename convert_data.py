import pandas as pd
import os
import sys

def convert_data_to_format(input_file, output_file=None):
    """
    Convert a CSV file to the required format for the card generator.
    
    Args:
        input_file (str): Path to the input CSV file
        output_file (str, optional): Path to the output CSV file. If None, the output
                                    will be saved as 'converted_data.csv'
    """
    try:
        # Read the input CSV file
        df = pd.read_csv(input_file)
        
        # Check if the required columns already exist
        required_columns = ['LAABHARTHI_NAME', 'CONTACT_NUMBER', 'ARPIT_GROUP', 'AREA']
        existing_columns = list(df.columns)
        
        # Map columns if different names are used
        columns_mapping = {}
        
        for col in required_columns:
            # Check if column already exists (case-insensitive comparison)
            if col in existing_columns:
                continue
                
            # Try to find a similar column name
            similar_cols = [ecol for ecol in existing_columns if 
                           col.lower() in ecol.lower() or 
                           ecol.lower() in col.lower()]
            
            # Try to guess columns based on common patterns
            if not similar_cols:
                if 'NAME' in col:
                    name_candidates = [c for c in existing_columns if 
                                      'name' in c.lower() or 
                                      'person' in c.lower() or
                                      'beneficiary' in c.lower()]
                    if name_candidates:
                        similar_cols = [name_candidates[0]]
                        
                elif 'CONTACT' in col or 'NUMBER' in col:
                    contact_candidates = [c for c in existing_columns if 
                                         'phone' in c.lower() or 
                                         'contact' in c.lower() or
                                         'mobile' in c.lower() or
                                         'number' in c.lower()]
                    if contact_candidates:
                        similar_cols = [contact_candidates[0]]
                        
                elif 'GROUP' in col:
                    group_candidates = [c for c in existing_columns if 
                                      'group' in c.lower() or 
                                      'category' in c.lower() or
                                      'team' in c.lower()]
                    if group_candidates:
                        similar_cols = [group_candidates[0]]
                        
                elif 'AREA' in col:
                    area_candidates = [c for c in existing_columns if 
                                      'area' in c.lower() or 
                                      'location' in c.lower() or
                                      'place' in c.lower() or
                                      'region' in c.lower() or
                                      'city' in c.lower()]
                    if area_candidates:
                        similar_cols = [area_candidates[0]]
            
            if similar_cols:
                columns_mapping[similar_cols[0]] = col
            else:
                print(f"Warning: Could not find a match for required column {col}")
                print(f"Available columns: {existing_columns}")
                user_input = input(f"Enter the column name to use for {col} (or leave blank to skip): ")
                if user_input.strip():
                    columns_mapping[user_input.strip()] = col
                else:
                    print(f"Skipping {col}. You'll need to add this column manually.")
        
        # Rename columns according to the mapping
        if columns_mapping:
            df = df.rename(columns=columns_mapping)
        
        # Check if all required columns exist after mapping
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            # Add empty columns for missing fields
            for col in missing_columns:
                df[col] = ""
            print(f"Warning: The following required columns were not found and are added as empty: {missing_columns}")
        
        # Ensure the output file path exists
        if output_file is None:
            output_file = 'converted_data.csv'
            
        # Save the converted data
        df.to_csv(output_file, index=False)
        print(f"Data converted successfully and saved to {output_file}")
        return True
    
    except Exception as e:
        print(f"Error converting data: {str(e)}")
        return False

if __name__ == "__main__":
    # Check if the script is called with an input file
    if len(sys.argv) < 2:
        print("Usage: python convert_data.py input_file.csv [output_file.csv]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_data_to_format(input_file, output_file) 