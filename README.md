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

AUTHOR
------
Carried

VERSION
-------
1.0 (January 10, 2025)
