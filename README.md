# Excel and PDF Data Scraper

This project demonstrates how to extract and consolidate data from Excel and PDF files into a structured format. It was created as an educational example to showcase data scraping and processing techniques.

## Project Overview

The project consists of two main components:

1. **Excel Scraper** (`excel-scraper.py`)
   - Processes Excel files from the `excel_files` directory
   - Extracts data from multiple sheets and consolidates it
   - Handles different data formats and categories
   - Updates a template Excel file with processed data

2. **PDF Scraper** (`pdf-scraper.py`)
   - Processes PDF files from the `pdf_files` directory
   - Extracts tabular data from PDFs using tabula-py
   - Filters data based on specific conditions
   - Consolidates data into a structured format
   - Updates a template Excel file with the processed data

## Features

- **Data Extraction**: Automatically extracts data from both Excel and PDF files
- **Data Processing**: Handles different data formats and structures
- **Template Integration**: Updates a template Excel file with processed data
- **Error Handling**: Includes basic error handling for different file formats
- **Data Filtering**: Filters data based on specific conditions (e.g., budget exceeded, completion issues)

## Requirements

- Python 3.x
- pandas
- openpyxl
- tabula-py

## Usage

1. Place your Excel files in the `excel_files` directory
2. Place your PDF files in the `pdf_files` directory
3. Run the appropriate scraper:
   ```bash
   python excel-scraper.py
   # or
   python pdf-scraper.py
   ```
4. The processed data will be saved in the template Excel file


## Notes

- The scrapers expect specific file naming conventions and structures
- Make sure to have a template Excel file in the respective directories
- The code includes comments explaining the main processing steps
