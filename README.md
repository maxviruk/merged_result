# Merge Excel Files

This Python script merges multiple Excel files from the `./data` folder into a single output file. It:

- Skips the first 16 rows of each file
- Removes duplicate rows based on the `Employee ID` column (if present)
- Saves the merged output as `merged_result.xlsx` in the same folder as the script
- If the output file already exists, it appends the current date in DDMM format to the filename

## How to Use

1. Place all Excel files to merge in the `data/` folder.
2. Run the script:
   ```bash
   python merge_files.py
   ```
3. The result will be saved next to the script as `merged_result.xlsx` or `merged_result_DDMM.xlsx`.

## Project Structure

```
├── data/               # Folder with source Excel files
├── merged_result.xlsx  # Merged output file (auto-generated)
├── merge_files.py      # This script
├── .gitignore          # Git exclusions
```

## .gitignore

This repository ignores the `data/` folder and all generated output files to avoid uploading sensitive or large files.

## Requirements

- Python 3.x
- pandas

Install dependencies:

```bash
pip install pandas
```
