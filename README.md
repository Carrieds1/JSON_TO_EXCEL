==================================
JSON to Excel Extractor
==================================

A utility tool that converts complex JSON data into well-formatted Excel spreadsheets.

DESCRIPTION
-----------
This tool extracts JSON objects from files or text dumps and converts them into
easily readable Excel spreadsheets. It's especially useful for analyzing nested
JSON structures by flattening them into a tabular format with hierarchical column names.

FEATURES
--------
* Flattens nested JSON objects and arrays into a single Excel table
* Preserves hierarchical relationships with dot notation (user.name) and indexed arrays (items[0])
* Automatically adjusts column widths for optimal readability
* Extracts individual JSON objects from larger text files using pattern matching
* Handles deeply nested structures with full support for complex data

REQUIREMENTS
-----------
* Python 3.6 or higher
* pandas library
* xlsxwriter library

INSTALLATION
-----------
1. Install required packages:
   pip install pandas xlsxwriter

2. Configure the input/output paths in the script:
   - DEFAULT_INPUT_PATH: Path to your JSON input file
   - DEFAULT_OUTPUT_PATH: Where to save the Excel output

USAGE
-----
1. Basic usage:
   python json_to_excel.py

2. Import and use in your own scripts:
   from json_to_excel import extract_json_data
   extract_json_data('input.json', 'output.xlsx')

3. Extract specific JSON objects from text:
   extract_json_data('input.txt', 'output.xlsx', 
                    start_pattern='{"created_at":', 
                    end_pattern='}')

4. Customization options:
   - Edit the script constants to change how JSON is processed
   - For complete JSON files: leave start_pattern and end_pattern as None
   - For specific JSON objects: define patterns that match your data structure
   - Modify DICT_SEPARATOR and array formatting to suit your preferences
   - Adjust column formatting for better readability

CONFIGURATION
------------
The script contains several customizable constants:

* JSON_START_PATTERN: Pattern that marks the start of each JSON object
* JSON_END_PATTERN: Pattern that marks the end of each JSON object
* DICT_SEPARATOR: Character used between parent/child keys (default: '.')
* INPUT_ENCODING: Character encoding for input files (default: 'utf-8')
* COLUMN_PADDING: Extra space added to Excel column width (default: 2)

EXAMPLE OUTPUT
-------------
Input JSON:
{
  "user": {
    "name": "John",
    "addresses": [
      {"city": "New York", "type": "home"},
      {"city": "Boston", "type": "work"}
    ]
  }
}

Excel columns:
user.name | user.addresses[0].city | user.addresses[0].type | user.addresses[1].city | user.addresses[1].type
---------|------------------------|------------------------|------------------------|-----------------------
John     | New York               | home                   | Boston                 | work

DATA FORMATTING
--------------
All data is exported in UTF-8 string format by default. To convert string values
to numbers for analysis in Excel:

1. Select the column(s) you wish to convert
2. Navigate to the "Data" tab in the Excel ribbon
3. Click "Text to Columns"
4. Select "Delimited" and click "Next"
5. Uncheck all delimiter options and click "Next"
6. Under "Column data format", select "General" (for automatic type detection)
7. Click "Finish"

This process tells Excel to reinterpret the text values as numbers where
appropriate, allowing you to perform calculations and numerical analysis.

For date fields, you may need to use Excel's date conversion functions
after completing the steps above.

AUTHOR
------
Carried

VERSION
-------
1.0 (January 10, 2025)
