import pandas as pd
import os
from pathlib import Path

# =============================================================================
# CONFIGURATION SECTION - EDIT THESE PATHS FOR YOUR NEEDS
# =============================================================================

# INPUT FILE PATH
# Change this to the path of your CSV file that needs cleaning
INPUT_CSV_PATH = "C:\\Scripts\\Input\\FILEtoCLEAN.csv"

# OUTPUT DIRECTORY PATH  
# Change this to where you want the cleaned CSV files to be saved
# The script will automatically create this directory if it doesn't exist
OUTPUT_DIRECTORY = "C:\\Scripts\\Output\\"

# OUTPUT FILE PREFIX (optional)
# This will be added to the beginning of your output filename
# Example: "cleaned_" will turn "data.csv" into "cleaned_data.csv"
OUTPUT_PREFIX = "cleaned_"

# =============================================================================
# END CONFIGURATION SECTION
# =============================================================================

def clean_csv(input_path, output_path=None):
    """
    Remove columns from a CSV file that have only one unique value in the data rows.
    Headers are excluded from the uniqueness check - only data rows (row 2 onwards) are analyzed.
    
    Args:
        input_path (str): Path to the input CSV file
        output_path (str, optional): Path for the output file. If None, uses default output directory.
    
    Returns:
        str: Path to the cleaned output file
    """
    try:
        # Validate input file exists
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Set default output path if not provided
        if output_path is None:
            input_filename = Path(input_path).name
            output_filename = f"{OUTPUT_PREFIX}{input_filename}"
            output_path = os.path.join(OUTPUT_DIRECTORY, output_filename)
        
        # Ensure output directory exists
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Load the CSV file
        print(f"Loading CSV file: {input_path}")
        df = pd.read_csv(input_path, low_memory=False)
        
        if len(df) == 0:
            print("Warning: CSV file is empty")
            return output_path
        
        # Identify columns with only one unique value in data rows (excluding header)
        # We skip the first row (index 0) which contains headers
        data_rows = df.iloc[1:] if len(df) > 1 else df
        
        single_value_cols = []
        for col in df.columns:
            unique_count = data_rows[col].nunique(dropna=False)
            if unique_count <= 1:
                single_value_cols.append(col)
        
        # Drop those columns
        df_cleaned = df.drop(columns=single_value_cols)
        
        # Save the cleaned file
        df_cleaned.to_csv(output_path, index=False)
        
        # Print summary
        print(f"\nProcessing complete!")
        print(f"Original columns: {len(df.columns)}")
        print(f"Remaining columns: {len(df_cleaned.columns)}")
        print(f"Removed columns ({len(single_value_cols)}):")
        
        if single_value_cols:
            for col in single_value_cols:
                # Show what the single value was
                if len(data_rows) > 0:
                    sample_value = data_rows[col].iloc[0] if not data_rows[col].empty else "N/A"
                    print(f"  - {col} (value: {sample_value})")
                else:
                    print(f"  - {col}")
        else:
            print("  None - all columns had multiple unique values")
        
        print(f"\nCleaned file saved to: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"Error processing CSV file: {str(e)}")
        raise

# =============================================================================
# SCRIPT EXECUTION - NO NEED TO EDIT BELOW THIS LINE
# =============================================================================

if __name__ == "__main__":
    # The script will use the paths defined in the CONFIGURATION SECTION above
    print("=" * 60)
    print("CSV Column Cleaner Script")
    print("=" * 60)
    print(f"Input file: {INPUT_CSV_PATH}")
    print(f"Output directory: {OUTPUT_DIRECTORY}")
    print(f"Output prefix: '{OUTPUT_PREFIX}'")
    print("=" * 60)
    
    # Run the cleaning function with the configured paths
    clean_csv(INPUT_CSV_PATH)
