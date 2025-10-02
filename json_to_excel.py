import json
import os
import pandas as pd

# Constants 
INITIAL_POSITION = 0
NOT_FOUND = -1  
LEGACY_RESULT_LENGTH = 3 
EXCEL_WRITER = 'xlsxwriter'
READ = 'r'

# JSON processing
DEFAULT_INPUT_PATH = "PATHNAME/FILENAME"
DEFAULT_OUTPUT_PATH = "PATHNAME/FILENAME"
JSON_START_PATTERN = '{"EXAMPLE": "'        # eg'{"created_at": "'  
JSON_END_PATTERN = '"EXAMPLE"}'             # '"lang": "und"}'   
INPUT_ENCODING = 'utf-8'   

# JSON patterns and separators   
DICT_SEPARATOR = '.'     
ARRAY_PREFIX = '['       
ARRAY_SUFFIX = ']'       
EMPTY_ARRAY = "[]"   
TRUNCATION_SEPERATOR = "..."     

# Excel formatting settings
COLUMN_PADDING = 2      
DEFAULT_SHEET_NAME = 'Data'
DEFAULT_VALUE = ""   

# Error and success messages
ERROR_INVALID_JSON = "Error: File is neither valid JSON nor contains extractable JSON objects"
ERROR_NO_DATA = "No data found to export"
SUCCESS_MESSAGE = "Data saved to {output_path} with auto-sized columns"
ERROR_JSON_PARSE = "Error parsing JSON at position {start_pos}-{end_pos}: {error_details}"
SNIPPET_LENGTH = 50                         

# Reads file content from the specified path with appropriate encoding
def read_file(input_path):
    input_file = os.path.expanduser(input_path)
    with open(input_file, READ, encoding=INPUT_ENCODING) as file_handle:
        return file_handle.read()

# Finds the start and end positions of a JSON object in content
def find_json_positions(file_content, start_pattern, current_position):
    start_pos = file_content.find(start_pattern, current_position)
    if start_pos == NOT_FOUND:
        return None, None
        
    end_pos = file_content.find(JSON_END_PATTERN, start_pos)
    if end_pos == NOT_FOUND:
        return None, None
    
    return start_pos, end_pos

# Extracts a JSON string from content and returns the new position for continued parsing
def extract_json_string(file_content, start_pos, end_pos):
    end_pos_with_pattern = end_pos + len(JSON_END_PATTERN)
    return file_content[start_pos:end_pos_with_pattern], end_pos_with_pattern

# Creates a truncated snippet of JSON string for error reporting
def create_error_snippet(json_str):
    if len(json_str) > SNIPPET_LENGTH:
        return json_str[:SNIPPET_LENGTH] + TRUNCATION_SEPERATOR
    else:
        return json_str

# Parses a JSON string and handles any parsing errors
def parse_json_string(json_str, start_pos, end_pos):
    try:
        return json.loads(json_str), None
    except json.JSONDecodeError as error:
        snippet = create_error_snippet(json_str)
        error_message = ERROR_JSON_PARSE.format(
            start_pos=start_pos, 
            end_pos=end_pos, 
            error_details=error
        )
        return None, (error_message, snippet)

# Processes a single JSON item from file content and returns object or error
def process_json_item(file_content, current_position, start_pattern):
    start_pos, end_pos = find_json_positions(file_content, start_pattern, current_position)
    if start_pos is None:
        return None, None, True, None  
    
    json_str, new_position = extract_json_string(file_content, start_pos, end_pos)
    
    json_obj, error = parse_json_string(json_str, start_pos, end_pos)
    
    return json_obj, error, False, new_position  

# Handles result from process_json_item with compatibility for different formats
def handle_process_result(result):
    if len(result) == LEGACY_RESULT_LENGTH:
        json_obj, error, stop_loop = result
        new_position = None
    else:
        json_obj, error, stop_loop, new_position = result
        
    return json_obj, error, stop_loop, new_position

# Processes and displays error information from JSON parsing
def process_error(error):
    if error:
        error_message, snippet = error
        print(error_message)
        print(f"Problem in string: {snippet}")

# Extracts multiple JSON objects from file content using pattern matching
def extract_multiple_objects(file_content, start_pattern, end_pattern):
    json_objects = []
    current_position = INITIAL_POSITION
    
    while True:
        result = process_json_item(file_content, current_position, start_pattern)
        
        json_obj, error, stop_loop, new_position = handle_process_result(result)
        
        if new_position is not None:
            current_position = new_position
            
        if stop_loop:
            break
            
        if json_obj:
            json_objects.append(json_obj)
        else:
            process_error(error)
    
    return json_objects

# Loads JSON from file content, handling both single objects and arrays
def load_json_file(file_content):
    try:
        parsed_json = json.loads(file_content)
        if not isinstance(parsed_json, list):
            return [parsed_json]
        else:
            return parsed_json
    except json.JSONDecodeError:
        return None

# Flattens a nested dictionary into a single-level dictionary with dot notation
def flatten_dict(nested_dict, parent_key, flattened_dict):
    for key, value in nested_dict.items():
        if parent_key:
            new_key = f"{parent_key}{DICT_SEPARATOR}{key}"
        else:
            new_key = key
        
        if isinstance(value, dict):
            flatten_dict(value, new_key, flattened_dict)
        elif isinstance(value, list):
            flatten_list(value, new_key, flattened_dict)
        else:
            flattened_dict[new_key] = value
    
    return flattened_dict

# Handles empty lists during flattening process
def handle_empty_list(nested_list, parent_key, flattened_dict):
    if not nested_list:
        flattened_dict[parent_key] = EMPTY_ARRAY
        return True
    return False

# Creates a key with array indexing notation for flattened JSON
def create_array_key(parent_key, index):
    return f"{parent_key}{ARRAY_PREFIX}{index}{ARRAY_SUFFIX}"

# Processes a list item based on its type during flattening
def process_list_item(item, array_key, flattened_dict):
    if isinstance(item, dict):
        flatten_dict(item, array_key, flattened_dict)
    elif isinstance(item, list):
        flatten_list(item, array_key, flattened_dict)
    else:
        flattened_dict[array_key] = item

# Flattens a nested list into a single-level dictionary with indexed notation
def flatten_list(nested_list, parent_key, flattened_dict):
    if handle_empty_list(nested_list, parent_key, flattened_dict):
        return flattened_dict
    
    for index, item in enumerate(nested_list):
        array_key = create_array_key(parent_key, index)
        process_list_item(item, array_key, flattened_dict)
    
    return flattened_dict

# Main function to flatten nested JSON data into a single-level dictionary
def flatten_json(nested_data, parent_key=DEFAULT_VALUE):
    flattened_dict = {}
    
    if isinstance(nested_data, dict):
        return flatten_dict(nested_data, parent_key, flattened_dict)
    elif isinstance(nested_data, list):
        return flatten_list(nested_data, parent_key, flattened_dict)
    else:
        flattened_dict[parent_key] = nested_data
        return flattened_dict

# Converts a value to its string representation or default value
def stringify_value(value):
    if value is not None:
        return str(value)
    else:
        return DEFAULT_VALUE

# Converts JSON records to Excel-compatible rows with string values
def convert_to_excel_rows(json_records):
    excel_rows = []
    
    for record in json_records:
        flat_record = flatten_json(record)
        string_record = {key: stringify_value(value) for key, value in flat_record.items()}
        excel_rows.append(string_record)
    
    return excel_rows

# Creates a pandas DataFrame from Excel rows, handling empty data
def create_dataframe(excel_rows):
    if not excel_rows:
        print(ERROR_NO_DATA)
        return None
    
    return pd.DataFrame(excel_rows)

# Calculates optimal column width for Excel output based on content
def calculate_column_width(data_frame, column_name):
    max_data_width = data_frame[column_name].astype(str).map(len).max()
    return max(max_data_width, len(str(column_name))) + COLUMN_PADDING

# Formats Excel columns with appropriate widths for readability
def format_excel_columns(data_frame, excel_writer):
    workbook = excel_writer.book
    worksheet = excel_writer.sheets[DEFAULT_SHEET_NAME]
    
    for col_index, column_name in enumerate(data_frame.columns):
        column_width = calculate_column_width(data_frame, column_name)
        worksheet.set_column(col_index, col_index, column_width)

# Sets up Excel writer with proper formatting and displays success message
def setup_excel_writer(data_frame, output_path):
    with pd.ExcelWriter(output_path, engine=EXCEL_WRITER) as excel_writer:
        data_frame.to_excel(excel_writer, index=False, sheet_name=DEFAULT_SHEET_NAME)
        format_excel_columns(data_frame, excel_writer)
    
    print(SUCCESS_MESSAGE.format(output_path=output_path))
    return True

# Writes Excel rows to an Excel file with appropriate formatting
def write_excel_file(excel_rows, output_path):
    data_frame = create_dataframe(excel_rows)
    if data_frame is not None:
        return setup_excel_writer(data_frame, output_path)
    
    return False

# Main function to extract JSON data from input file and save to Excel
def extract_json_data(input_path, output_path, start_pattern=None, end_pattern=None):
    file_content = read_file(input_path)
    
    if start_pattern and end_pattern:
        json_records = extract_multiple_objects(file_content, start_pattern, end_pattern)
    else:
        json_records = load_json_file(file_content)
        if json_records is None:
            print(ERROR_INVALID_JSON)
            return
    
    excel_rows = convert_to_excel_rows(json_records)
    write_excel_file(excel_rows, output_path)

if __name__ == "__main__":
    extract_json_data(
        DEFAULT_INPUT_PATH, 
        DEFAULT_OUTPUT_PATH,
        JSON_START_PATTERN,
        JSON_END_PATTERN
    )
