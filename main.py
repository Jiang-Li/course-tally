import os
import sys
import logging
import pandas as pd
from datetime import datetime
from src.tally_reader import read_tally_file
from src.course_scraper import CourseScraperFranklin
from src.excel_handler import update_excel_sheet

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

def setup_directories():
    """Create necessary directories if they don't exist."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    dirs = ['data', 'logs']
    for dir_name in dirs:
        dir_path = os.path.join(current_dir, dir_name)
        os.makedirs(dir_path, exist_ok=True)
    return current_dir

def process_course_details(scraper, course_code, term, output_dir):
    """Process and save course details."""
    try:
        course_details = scraper.get_course_details(
            course_code=course_code,
            term=term
        )
        
        if course_details['status'] == 'Found':
            # Create filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{course_code}_{term.replace(' ', '_')}_{timestamp}.txt"
            output_file = os.path.join(output_dir, filename)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"Course: {course_code}\n")
                f.write(f"Term: {term}\n")
                f.write(f"Retrieved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("\nDescription:\n")
                f.write(course_details['description'])
            
            logging.info(f"Course description saved to {output_file}")
            return True
        else:
            logging.error(f"Error: {course_details['status']}")
            return False
            
    except Exception as e:
        logging.error(f"Error processing course {course_code}: {str(e)}")
        return False

def main():
    try:
        # Setup directories
        current_dir = setup_directories()
        data_dir = os.path.join(current_dir, 'data')
        
        # Initialize scraper
        scraper = CourseScraperFranklin(headless=True)
        
        try:
            # Read tally file
            tally_file = "Course Tally.xlsx"
            table, is_unique, key_cols = read_tally_file(tally_file)
            
            logging.info("\nFound table with 'Subj':")
            logging.info(f"Shape: {table.shape}")
            logging.info(f"Columns: {list(table.columns)}")
            logging.info(f"Rows are {'uniquely' if is_unique else 'not uniquely'} identified by {', '.join(key_cols)}")
            
            # Process each course in the tally
            if 'Subj' in table.columns and 'Crse' in table.columns:
                for _, row in table.iterrows():
                    course_code = f"{row['Subj']}{row['Crse']}"
                    process_course_details(scraper, course_code, "Spring 2025", data_dir)
            else:
                logging.error("Required columns 'Subj' and 'Crse' not found in tally file")
                
        except Exception as e:
            logging.error(f"Error reading tally file: {str(e)}")
            
    except Exception as e:
        logging.error(f"Critical error: {str(e)}")
        
    finally:
        # Cleanup
        try:
            scraper.close_browser()
        except:
            pass

if __name__ == "__main__":
    main() 