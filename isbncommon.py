from pathlib import Path

def find_desired_val_from_search_val(search_val, ws, ws_search_val_col_index, ws_desired_val_col_index):
    """Returns the desired value from a `Worksheet` given a unique, non-empty search value and location details to match.

    Keyword arguments:
    search_val -- value to find in the Workbook
    ws -- `Worksheet` containing values to search
    ws_search_val_col_index -- one-based index of the column in the `Worksheet` containing search value
    ws_desired_val_col_index -- one-based index of the column in the `Worksheet` containing desired value
    """
    # Note: `row` is a tuple of cells, i.e., `row` has a zero-based index
    #        each `cell` in the row, however, has a one-based index 
    for row in ws.iter_rows(min_row = 2):
        val = row[ws_search_val_col_index - 1].value
        
        # stop searching if search val is empty
        if(search_val is None or search_val == ''):
            return None
        
        if(search_val is not None and val is not None):
            search_val_type = type(val)
            val_type = type(val)

            # cast to string if both values are not of the same type
            if(search_val_type != val_type):
                search_val = str(val)
                val = str(search_val)
            
            # remove whitespace for string values
            if(search_val_type == str):
                search_val = search_val.strip()
                val = val.strip()
            
            # check if value is what we're looking for
            if(search_val == val):
                desired_val = row[ws_desired_val_col_index - 1].value
                return desired_val
            
    return None

def find_row_from_title(title_to_find, ws, ws_title_col_index):
    for row in ws.rows:
        title = row[ws_title_col_index - 1].value
        
        if(title == title_to_find):
            return row
    return None

def find_row_from_search_val(search_val, ws, search_val_col_index):
    for row in ws.rows:
        title = row[search_val_col_index - 1].value
        
        if(title == search_val):
            return row
    return None
   
def get_xlsx_files(dir_path):
    """Returns the list of .xlsx files as `Path` objects in a directory

    Keyword arguments:
    dir_path -- the directory to find files
    """
    p = Path(dir_path)
    files = list(p.glob('*.xlsx'))
    return files

def print_worksheet_names(wb):
    """Print the names of all Worksheets in a given Workbook.

    Keyword arguments: 
    wb -- the Workbook containing desired Worksheet names to print
    """
    for ws in wb.worksheets:
        print(ws.title)

def print_worksheet_header_columns(ws):
    """Prints the value and coordinates of all header columns in a Worksheet.

    Keyword arguments: 
    ws -- the Worksheet containing desired column headers to print
    """
    for row in ws.iter_rows(max_row = 1):
        for cell in row:
            print(cell.coordinate, cell.value)
