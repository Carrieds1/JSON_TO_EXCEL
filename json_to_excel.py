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
# will ensure portability across multiple systems
import os
import pandas as pd

NOT_FOUND = -1  # Return value when pattern not found
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


"""
Recursively flatten nested dictionaries and lists into a single-level dictionary.
Parent-child relationships are preserved in the key names.

Args:
    nested_data: The nested dictionary or list to flatten
    parent_key: The base key to prepend to all keys in this structure
    
Returns:
    A dictionary with dot notation for nested objects
"""
def flatten_json(nested_data, parent_key=DEFAULT_VALUE):
    flattened_dict = {}
    
    # Handle dictionaries
    if isinstance(nested_data, dict):
        for key, value in nested_data.items():
            # Create hierarchical key name
            if parent_key:
                new_key = f"{parent_key}{DICT_SEPARATOR}{key}"
            else:
                new_key = key
            
            # Recursively flatten nested structures
            if isinstance(value, (dict, list)):
                flattened_dict.update(flatten_json(value, new_key))
            else:
                flattened_dict[new_key] = value
                
    # Handle lists
    elif isinstance(nested_data, list):
        if not nested_data:  # Empty list
            flattened_dict[parent_key] = EMPTY_ARRAY
        else:
            for index, item in enumerate(nested_data):
                # Create indexed key for array elements
                array_key = f"{parent_key}{ARRAY_PREFIX}{index}{ARRAY_SUFFIX}"
                
                # Recursively flatten nested structures
                if isinstance(item, (dict, list)):
                    flattened_dict.update(flatten_json(item, array_key))
                else:
                    flattened_dict[array_key] = item
    else:
        # Base case: primitive value
        flattened_dict[parent_key] = nested_data
        
    return flattened_dict

"""
Extract multiple JSON objects from text content using patterns.

Args:
    file_content: The raw content to search through
    start_pattern: Text pattern that marks the start of each JSON object
    end_pattern: Text pattern that marks the end of each JSON object
    
Returns:
    List of parsed JSON objects
"""
def extract_multiple_objects(file_content, start_pattern, end_pattern):
    json_objects = []
    current_position = 0
    
    while True:
        # Find boundaries of next JSON object
        start_pos = file_content.find(start_pattern, current_position)
        if start_pos == NOT_FOUND:
            break
            
        end_pos = file_content.find(end_pattern, start_pos)
        if end_pos == NOT_FOUND:
            break
        
        # Extract the complete JSON string (including the end pattern)
        end_pos_with_pattern = end_pos + len(end_pattern)
        json_str = file_content[start_pos:end_pos_with_pattern]
        
        try:
            # Parse the JSON object
            json_obj = json.loads(json_str)
            json_objects.append(json_obj)
        except json.JSONDecodeError as e:
            # Get a short snippet of the problematic JSON (first 50 chars is default)
            snippet = json_str[:SNIPPET_LENGTH] + "..." if len(json_str) > SNIPPET_LENGTH else json_str
            # Print error with position information and snippet
            print(ERROR_JSON_PARSE.format(start_pos, end_pos, e))
            print(f"Problem in string: {snippet}")
            # Skip invalid JSON without crashing
            pass
        
        # Move to position after the current JSON object
        current_position = end_pos_with_pattern
    
    return json_objects

"""
Extract JSON objects from a file and save them to Excel with horizontal layout.

Args:
    input_path: Path to the input file containing JSON objects
    output_path: Path to save the Excel output
    start_pattern: Pattern that marks the start of each JSON object (optional)
    end_pattern: Pattern that marks the end of each JSON object (optional)
"""
def extract_json_data(input_path, output_path, start_pattern=None, end_pattern=None):
    # Ensure paths are valid across platforms (windows, linux, mac)
    input_file = os.path.expanduser(input_path)
    
    # Read the input file
    with open(input_file, READ, encoding=INPUT_ENCODING) as file_handle:
        file_content = file_handle.read()
    
    # Determine method to extract JSON objects based on input format
    if start_pattern and end_pattern:
        # Extract multiple JSON objects using patterns (for text files with embedded JSON)
        json_records = extract_multiple_objects(file_content, start_pattern, end_pattern)
    else:
        # Try to parse as a clean JSON file (either array or single object)
        try:
            parsed_json = json.loads(file_content)
            if isinstance(parsed_json, list):
                json_records = parsed_json  # Already a list of objects
            else:
                json_records = [parsed_json]  # Single object wrapped in list
        except json.JSONDecodeError:
            print(ERROR_INVALID_JSON)
            return
    
    # Prepare data for Excel export
    excel_rows = []
    
    # Process each JSON record
    for record in json_records:
        # Flatten the entire object structure to a single level
        flat_record = flatten_json(record)
        
        # Convert all values to strings for Excel compatibility
        string_record = {}
        for key, value in flat_record.items():
            if value is not None:
                string_record[key] = str(value)
            else:
                string_record[key] = DEFAULT_VALUE

        excel_rows.append(string_record)
    
    # Create DataFrame and save to Excel with auto-sized columns
    if excel_rows:
        # Convert to pandas DataFrame
        df = pd.DataFrame(excel_rows)
        
        # Create Excel file with auto-sized columns
        with pd.ExcelWriter(output_path, engine=EXCEL_WRITER) as excel_writer:
            # Write the main data
            df.to_excel(excel_writer, index=False, sheet_name=DEFAULT_SHEET_NAME)
            
            # Access Excel objects for formatting
            workbook = excel_writer.book
            worksheet = excel_writer.sheets[DEFAULT_SHEET_NAME]
            
            # Auto-adjust column widths based on content
            for col_index, column_name in enumerate(df.columns):
                # Calculate ideal column width + padding for readability
                column_width = max(
                                   # Maximum data width in this column
                                   df[column_name].astype(str).map(len).max(),
                                   # Header width
                                   len(str(column_name)))
                column_width += COLUMN_PADDING
                
                # Apply the calculated width
                worksheet.set_column(col_index, col_index, column_width)

        print(SUCCESS_MESSAGE.format(output_path))
    else:
        print(ERROR_NO_DATA)

# Code to run when file is executed directly
if __name__ == "__main__":
    # Run the extraction with JSON-specific settings
    extract_json_data(
        DEFAULT_INPUT_PATH, 
        DEFAULT_OUTPUT_PATH,
        JSON_START_PATTERN,
        JSON_END_PATTERN
    )