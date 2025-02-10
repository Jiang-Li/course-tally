# Course Tally Reader

A Python tool for reading, processing, and analyzing course tally data from Excel files. This tool helps in managing course information, including scraping course details and maintaining organized records.

## Features

- Read and process course tally data from Excel files
- Automatically detect and parse tables with course information
- Check for duplicate entries and data integrity
- Scrape course details from Franklin University's system
- Export course descriptions and details to text files
- Update Excel sheets with course information

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/course-tally-reader.git
cd course-tally-reader
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
2. Process and validate the data
3. Display summary information about the courses

### Advanced Usage

The tool provides several modules that can be used independently:

- `tally_reader.py`: Core functionality for reading and processing tally files
- `excel_handler.py`: Functions for Excel file operations
- `course_scraper.py`: Web scraping functionality for course details

## Project Structure

```
course-tally-reader/
├── main.py              # Main entry point
├── requirements.txt     # Project dependencies
├── data/               # Directory for output files
└── src/
    ├── tally_reader.py  # Core tally reading functionality
    ├── excel_handler.py # Excel processing functions
    └── course_scraper.py# Course scraping functionality
```

## Dependencies

- Python 3.8+
- pandas >= 1.5.0
- openpyxl >= 3.0.0
- selenium (for web scraping)

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details. 