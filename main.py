import os
import sys
import logging
import pandas as pd
from src.tally_reader import read_tally_file

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

def main():
    try:
        # Read tally file
        tally_file = "Course Tally.xlsx"
        table, _, _ = read_tally_file(tally_file)
        
        # Display shape
        logging.info(f"DataFrame shape: {table.shape}")
        
        # Check for duplicates
        duplicates = table.duplicated().sum()
        if duplicates > 0:
            logging.info(f"Found {duplicates} duplicate rows")
            # Show the duplicate rows
            duplicate_rows = table[table.duplicated(keep=False)]
            logging.info("\nDuplicate rows:")
            logging.info(duplicate_rows)
        else:
            logging.info("No duplicate rows found")

    except Exception as e:
        logging.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main() 