# Course Tally Reader

A Python tool for reading and analyzing course tally data from Excel files. This tool helps in analyzing course information by checking for data shape and duplicates.

## Features
- Read and process course tally data from Excel files (Course Tally.xlsx), which can be downloaded from the course tally BI dashboard
- Compare and update course information between tally data and Leeds course data
- Display column mapping between files to identify matched and unmatched fields
- Clean and standardize data fields like course numbers and days for consistent comparison
- Track and report any unmatched courses or data discrepancies

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Jiang-Li/course-tally.git
cd course-tally
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```python
python main.py
```

This will:
1. Read the course tally Excel file
2. Display the shape of the data (rows × columns)
3. Check and report any duplicate entries

## Project Structure

```
course-tally/
├── main.py              # Main entry point
├── requirements.txt     # Project dependencies
└── src/
    └── tally_reader.py  # Core tally reading functionality
```

## Dependencies

- Python 3.8+
- pandas
- openpyxl

## License

This project is licensed under the MIT License - see the LICENSE file for details. 