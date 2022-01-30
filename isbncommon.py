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

def find_isbn_from_barcode(barcode_to_find, ws, ws_barcode_col_index, ws_isbn_col_index):
    """Returns the ISBN from a Worksheet given a barcode and location details to match.

    Keyword arguments:
    barcode_to_find -- barcode string to find in the Workbook
    ws -- Worksheet containing ISBNs and barcodes to search
    ws_barcode_col_index -- one-based index of the column in the Worksheet containing barcode data
    ws_isbn_col_index -- one-based index of the column in the Worksheet containing ISBN data
    """
    # Note: `row` is a tuple of cells, i.e., `row` has a zero-based index
    #        each `cell` in the row, however, has a one-based index 
    for row in ws.rows:
        barcode = row[ws_barcode_col_index - 1].value
        
        if(barcode == barcode_to_find):
            isbn = row[ws_isbn_col_index - 1].value
            return isbn

    return ''

   