import pandas as pd
from typing import List
import os
from src.course_scraper import CourseScraperFranklin

def update_excel_sheet(term: str, course_names: List[str], excel_path: str) -> None:
    """
    Create or update an Excel file with course names in a sheet named after the term.
    
    Args:
        term (str): Term name (e.g., Fall 2024)
        course_names (List[str]): List of course names
        excel_path (str): Path to the Excel file
    """
    try:
        # Create DataFrame with all required columns
        df = pd.DataFrame(columns=[
            'course_name',
            'time',
            'date',
            'location',
            'classroom',
            'instructor',
            'status'
        ])
        
        # Initialize course scraper
        scraper = CourseScraperFranklin()
        
        # Get details for each course
        for course in course_names:
            details = scraper.get_course_details(course, term)
            if details:
                df = pd.concat([df, pd.DataFrame([{
                    'course_name': course,
                    'time': details.get('time', 'N/A'),
                    'date': details.get('date', 'N/A'),
                    'location': details.get('location', 'N/A'),
                    'classroom': details.get('classroom', 'N/A'),
                    'instructor': details.get('instructor', 'N/A'),
                    'status': details.get('status', 'Error')
                }])], ignore_index=True)
        
        # If file exists, load it and update/create sheet
        if os.path.exists(excel_path):
            with pd.ExcelWriter(excel_path, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=term, index=False)
        else:
            # Create new Excel file with the sheet
            df.to_excel(excel_path, sheet_name=term, index=False)
            
    except Exception as e:
        raise Exception(f"Error updating Excel file: {str(e)}") 