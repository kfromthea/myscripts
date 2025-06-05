###this script converts excel files to pipe delimited files. just insert the path for the report_export_Table_1
    # saved to your computer and the script will save a copy of the pipe delimited version in your downloads folder.
    #libraries to download: openpyxl, pandas

import os
import pandas as pd
from datetime import datetime
import warnings

def convert_xlsm_to_pipe_delimited():
    # Suppress the data validation warning
    warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

    # Prompt user for file path and clean it
    xlsm_file_path = input("Enter the path to the macro-enabled Excel file (.xlsm): ").strip().strip('"').strip("'")

    # Get today's date in YYYY-MM-DD format
    today = datetime.today().strftime('%Y%m%d')

    # Generate output file name with date
    base, _ = os.path.splitext(xlsm_file_path)
    output_file_path = f"{base}_pipe_{today}.txt"

    try:
        # Load the workbook and read the first sheet
        df = pd.read_excel(xlsm_file_path, engine='openpyxl')  # Add sheet_name if needed

        # Save to pipe-delimited text
        df.to_csv(output_file_path, sep='|', index=False)

        print(f"\n✅ Conversion complete! Pipe-delimited file saved as:\n{output_file_path}")

    except FileNotFoundError:
        print("\n❌ File not found. Please check the path and try again.")
    except Exception as e:
        print(f"\n❌ An error occurred: {e}")

# Run the function
convert_xlsm_to_pipe_delimited()
