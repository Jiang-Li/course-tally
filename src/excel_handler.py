import pandas as pd
from typing import List, Dict, Optional
import os
from pathlib import Path

class ExcelFileError(Exception):
    """Custom exception for Excel file operations."""
    pass

def find_excel_file(exclude_filename: str, directory: str = ".") -> str:
    """
    Find an Excel file in the specified directory, excluding a specific file.
    
    Args:
        exclude_filename (str): Name of the file to exclude
        directory (str): Directory to search in (default: current directory)
        
    Returns:
        str: Path to the found Excel file
        
    Raises:
        ExcelFileError: If no suitable Excel file is found or multiple files exist
    """
    excel_extensions = ('.xlsx', '.xls')
    excel_files = [
        f for f in os.listdir(directory) 
        if f.endswith(excel_extensions) and f != exclude_filename
    ]
    
    if not excel_files:
        raise ExcelFileError(f"No Excel files found in {directory} other than {exclude_filename}")
    
    if len(excel_files) > 1:
        raise ExcelFileError(
            f"Multiple Excel files found: {', '.join(excel_files)}. "
            "Please specify which file to use."
        )
        
    return os.path.join(directory, excel_files[0])

def read_excel_until_empty(file_path: str, sheet_name: int = 0) -> pd.DataFrame:
    """
    Read an Excel file until the first empty row is encountered.
    
    Args:
        file_path (str): Path to the Excel file
        sheet_name (int): Index of the sheet to read (default: 0)
        
    Returns:
        pd.DataFrame: DataFrame containing the data
        
    Raises:
        ExcelFileError: If there's an error reading the file
    """
    if not os.path.exists(file_path):
        raise ExcelFileError(f"File not found: {file_path}")
            
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        
        # Find the first empty row
        first_empty_row = None
        for idx, row in df.iterrows():
            if row.isna().all():
                first_empty_row = idx
                break
        
        # Truncate DataFrame at first empty row if found
        if first_empty_row is not None:
            df = df.iloc[:first_empty_row]
        
        return df
        
    except Exception as e:
        raise ExcelFileError(f"Error reading Excel file {file_path}: {str(e)}")

# Alias functions for backward compatibility
find_leeds_file = find_excel_file
read_leeds_courses = read_excel_until_empty 