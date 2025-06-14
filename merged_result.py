import os
import pandas as pd

# Specify the path to the folder containing the Excel files
folder_path = r 'C:\Users\VIRIKMA\OneDrive - Anheuser-Busch InBev\My Documents\Мои полученные файлы\7. EUR - LoA\5. Analytics\1. 1Y plan\Data\Employee Master Data Workday'

# Get a list of all Excel files in the folder
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# List to store all dataframes
all_dataframes = []

for file in files:
    file_path = os.path.join(folder_path, file)
    
    # Read the Excel file, skipping the first 16 rows
    df = pd.read_excel(file_path, skiprows=16)
    
    all_dataframes.append(df)

# Combine all dataframes into one
combined_df = pd.concat(all_dataframes, ignore_index=True)

# Remove duplicates based on the 'Employee ID' column
if 'Employee ID' in combined_df.columns:
    combined_df.drop_duplicates(subset='Employee ID', keep='first', inplace=True)
else:
    print("⚠️ Warning: 'Employee ID' column not found in the files.")

# Save the final merged dataframe to an Excel file
output_path = os.path.join(folder_path, 'merged_result.xlsx')
combined_df.to_excel(output_path, index=False)

print(f"✅ Done! File saved as: {output_path}")
