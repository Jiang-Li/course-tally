import os
from pathlib import Path
import pandas as pd
from src.tally_reader import read_tally_file
from src.excel_handler import find_leeds_file, read_leeds_courses, ExcelFileError

def analyze_duplicates(df: pd.DataFrame) -> tuple[bool, pd.DataFrame]:
    """
    Analyze duplicates in a DataFrame.
    
    Args:
        df (pd.DataFrame): DataFrame to analyze
        
    Returns:
        tuple[bool, pd.DataFrame]: (has_duplicates, duplicate_rows)
    """
    duplicates = df.duplicated()
    duplicate_rows = df[df.duplicated(keep=False)] if duplicates.sum() > 0 else pd.DataFrame()
    return bool(duplicates.sum()), duplicate_rows

def main():
    try:
        # Read tally file
        tally_file = "Course Tally.xlsx"
        table, _, _ = read_tally_file(tally_file)
        print(f"DataFrame shape: {table.shape}")
        
        # Check for duplicates
        has_duplicates, duplicate_rows = analyze_duplicates(table)
        if has_duplicates:
            print(f"Found {len(duplicate_rows)} duplicate rows:")
            print(duplicate_rows)
        else:
            print("No duplicate rows found")
            
        # Find and read Leeds courses file
        try:
            leeds_file = find_leeds_file(tally_file)
            print(f"Found Leeds courses file: {leeds_file}")
            leeds_table = read_leeds_courses(leeds_file)
            print(f"Leeds courses DataFrame shape: {leeds_table.shape}")
        except ExcelFileError as e:
            print(f"Error processing Leeds file: {str(e)}")
            return

    except Exception as e:
        print(f"Error: {str(e)}")
        return

if __name__ == "__main__":
    main() 