"""
JSON to Excel Extractor

Extracts JSON data from files or text dumps and formats it into well-structured Excel spreadsheets.
Handles deeply nested JSON objects with hierarchical column names for easy data analysis.

Features:
- Flattens complex nested JSON structures
- Auto-sizes Excel columns for readability
- Customizable JSON extraction patterns
- Full support for arrays and nested dictionaries

Usage:
    python extract_monster.py

Dependencies:
    pandas, xlsxwriter

Author: [Carried]
Version: 1.0
Date: [1/10/25]

TODO: Add support for HTTP responses and direct URL input
"""

import json
import os
import pandas as pd

NOT_FOUND = -1  
EXCEL_WRITER = 'xlsxwriter'
READ = 'r'

# Below are edittable definitions for portability with other JSON stringdumps
# File handling constants
DEFAULT_INPUT_PATH = "PATHNAME/FILENAME"
DEFAULT_OUTPUT_PATH = "PATHNAME/FILENAME"

# Encoding constants
INPUT_ENCODING = 'utf-8'   # Encoding for reading input files
OUTPUT_ENCODING = 'utf-8'  # Encoding for writing output files - removed usage for excel output

# JSON parsing constants - change patterns to identify start and end of each JSON
JSON_START_PATTERN = '{"created_at": "'  # Start marker for each JSON object
JSON_END_PATTERN = '"lang": "und"}'      # End marker for each JSON object

# Key formatting constants
DICT_SEPARATOR = '.'     # Symbol used between parent/child keys (e.g., user.name)
ARRAY_PREFIX = '['       # Opening bracket for array indices
ARRAY_SUFFIX = ']'       # Closing bracket for array indices
EMPTY_ARRAY = "[]"       # Representation for empty arrays

# Excel formatting constants
COLUMN_PADDING = 2       # Extra space added to column width
DEFAULT_SHEET_NAME = 'Data'
DEFAULT_VALUE = ""       # Value to use for missing data

# Error messages
ERROR_INVALID_JSON = "Error: File is neither valid JSON nor contains extractable JSON objects"
ERROR_NO_DATA = "No data found to export"
SUCCESS_MESSAGE = "Data saved to {0} with auto-sized columns"
ERROR_JSON_PARSE = "Error parsing JSON at position {0}-{1}: {2}"
SNIPPET_LENGTH = 50  # Length of ERROR JSON PARSE preview to show in error messages

def flatten_json(nested_data, parent_key=DEFAULT_VALUE):
    flattened_dict = {}
    
    if isinstance(nested_data, dict):
        for key, value in nested_data.items():
            if parent_key:
                new_key = f"{parent_key}{DICT_SEPARATOR}{key}"
            else:
                new_key = key
            
            if isinstance(value, (dict, list)):
                flattened_dict.update(flatten_json(value, new_key))
            else:
                flattened_dict[new_key] = value
                
    elif isinstance(nested_data, list):
        if not nested_data: 
            flattened_dict[parent_key] = EMPTY_ARRAY
        else:
            for index, item in enumerate(nested_data):
                array_key = f"{parent_key}{ARRAY_PREFIX}{index}{ARRAY_SUFFIX}"
                
                if isinstance(item, (dict, list)):
                    flattened_dict.update(flatten_json(item, array_key))
                else:
                    flattened_dict[array_key] = item
    else:
        flattened_dict[parent_key] = nested_data
        
    return flattened_dict

def extract_multiple_objects(file_content, start_pattern, end_pattern):
    json_objects = []
    current_position = 0
    
    while True:
        start_pos = file_content.find(start_pattern, current_position)
        if start_pos == NOT_FOUND:
            break
            
        end_pos = file_content.find(end_pattern, start_pos)
        if end_pos == NOT_FOUND:
            break
        
        end_pos_with_pattern = end_pos + len(end_pattern)
        json_str = file_content[start_pos:end_pos_with_pattern]
        
        try:
            json_obj = json.loads(json_str)
            json_objects.append(json_obj)
        except json.JSONDecodeError as e:
            snippet = json_str[:SNIPPET_LENGTH] + "..." if len(json_str) > SNIPPET_LENGTH else json_str
            print(ERROR_JSON_PARSE.format(start_pos, end_pos, e))
            print(f"Problem in string: {snippet}")
            pass
        
        current_position = end_pos_with_pattern
    
    return json_objects

def extract_json_data(input_path, output_path, start_pattern=None, end_pattern=None):
    input_file = os.path.expanduser(input_path)
    
    with open(input_file, READ, encoding=INPUT_ENCODING) as file_handle:
        file_content = file_handle.read()
    
    if start_pattern and end_pattern:
        json_records = extract_multiple_objects(file_content, start_pattern, end_pattern)
    else:
        try:
            parsed_json = json.loads(file_content)
            if isinstance(parsed_json, list):
                json_records = parsed_json 
            else:
                json_records = [parsed_json]
        except json.JSONDecodeError:
            print(ERROR_INVALID_JSON)
            return
    
    excel_rows = []
    
    for record in json_records:
        flat_record = flatten_json(record)
        
        string_record = {}
        for key, value in flat_record.items():
            if value is not None:
                string_record[key] = str(value)
            else:
                string_record[key] = DEFAULT_VALUE

        excel_rows.append(string_record)
    
    if excel_rows:
        df = pd.DataFrame(excel_rows)
        
        with pd.ExcelWriter(output_path, engine=EXCEL_WRITER) as excel_writer:
            df.to_excel(excel_writer, index=False, sheet_name=DEFAULT_SHEET_NAME)
            
            workbook = excel_writer.book
            worksheet = excel_writer.sheets[DEFAULT_SHEET_NAME]
            
            for col_index, column_name in enumerate(df.columns):
                column_width = max(
                                   df[column_name].astype(str).map(len).max(),
                                   len(str(column_name)))
                column_width += COLUMN_PADDING
                
                worksheet.set_column(col_index, col_index, column_width)

        print(SUCCESS_MESSAGE.format(output_path))
    else:
        print(ERROR_NO_DATA)

if __name__ == "__main__":
    extract_json_data(
        DEFAULT_INPUT_PATH, 
        DEFAULT_OUTPUT_PATH,
        JSON_START_PATTERN,
        JSON_END_PATTERN
    )
