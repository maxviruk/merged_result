import os
import pandas as pd
from datetime import datetime

# --- CONFIGURABLE PATHS ---
script_dir = os.path.dirname(os.path.abspath(__file__))
source_folder = os.path.join(script_dir, 'data')  # Excel files should be in ./data/
output_folder = script_dir                        # Output will be in the same folder as the script

# Prepare output file name
base_filename = 'merged_result.xlsx'
output_file = os.path.join(output_folder, base_filename)

# Check if file already exists and create a new name with today's date if needed
if os.path.exists(output_file):
    date_suffix = datetime.now().strftime("%d%m")
    output_file = os.path.join(output_folder, f"merged_result_{date_suffix}.xlsx")

# Get list of Excel files
files = [f for f in os.listdir(source_folder) if f.endswith('.xlsx')]

print(f"📁 Found {len(files)} Excel files in source folder.")
all_dataframes = []

for file in files:
    file_path = os.path.join(source_folder, file)
    print(f"🔄 Reading: {file}")

    try:
        df = pd.read_excel(file_path, skiprows=16)

        if df.empty:
            print(f"⚠️ Skipped (empty after skipping 16 rows): {file}")
        else:
            all_dataframes.append(df)
    except Exception as e:
        print(f"❌ Error reading {file}: {e}")

# Check if there is data to merge
if not all_dataframes:
    print("🚫 No valid data to merge. Exiting.")
else:
    combined_df = pd.concat(all_dataframes, ignore_index=True)

    if 'Employee ID' in combined_df.columns:
        combined_df.drop_duplicates(subset='Employee ID', keep='first', inplace=True)
        print("✅ Duplicates removed based on 'Employee ID'.")
    else:
        print("⚠️ 'Employee ID' column not found. Skipping de-duplication.")

    # Save result
    combined_df.to_excel(output_file, index=False)
    print(f"✅ Done! Merged file saved to:\n{output_file}")
