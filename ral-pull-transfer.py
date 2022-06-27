from fileinput import filename
import os
from time import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from isbncommon import *

CALLNUMBER_COL_FACULTY = None
BARCODE_COL_FACULTY = None
ISBN_COL_FACULTY = None
TRANSFERNAME_COL_FACULTY = None

CALLNUMBER_COL_TRANSFER = column_index_from_string("A")
TITLE_COL_TRANSFER = column_index_from_string("B")
ISBN_COL_TRANSFER = column_index_from_string("C")
BARCODE_COL_TRANSFER = column_index_from_string("D")
WORLDCAT_COL_TRANSFER = column_index_from_string("E")
FACULTY_COL_TRANSFER = column_index_from_string("F")

lib_guide_dir = Path(os.getcwd(), 'input', 'lib_guide_title_lists')
faculty_keep_dir = Path(
    os.getcwd(), 'input', 'faculty_keep_lists', 'faculty_keep_masterlist.xlsx'
    )
pull_withdraw_list_dir = Path(os.getcwd(), 'output', 'pull_withdraw_lists')

# master list of all faculty requests compiled via email
faculty_keep_wb = load_workbook(Path(faculty_keep_dir))
faculty_keep_ws = faculty_keep_wb.worksheets[0]


def main():
    
    # master list of all faculty requests found in the withdrawal lists that will be transfered to their offices
    pull_transfer_wb = openpyxl.Workbook()
    pull_transfer_ws = pull_transfer_wb.active

    # set up header row into "pull & transfer" master list
    pull_transfer_ws.append(['Call Number', 'Title', 'ISBN', 'Barcode', 'Worldcat OCLC Number', 'Faculty'])

    # get the columns
    for row in faculty_keep_ws.iter_rows(max_row = 1):
        for cell in row:
            col_header = cell.value
            if(col_header is not None):
                col_header = col_header.lower()
                if(col_header == 'faculty'):
                    TRANSFERNAME_COL_FACULTY = cell.column
                elif(col_header == 'barcode'):
                    BARCODE_COL_FACULTY = cell.column

    # stop execution if required columns are not present
    if(TRANSFERNAME_COL_FACULTY is None or TRANSFERNAME_COL_FACULTY == ''):
        print('Error: Keep masterlist file is missing Faculty column.')
        return
    elif(BARCODE_COL_FACULTY is None or BARCODE_COL_FACULTY == ''):
        print('Error: Keep masterlist is missing Barcode column.')
        return



    # iterate through all categorized withdrawal files to create final "pull & withdraw" lists
    lib_guide_files = get_xlsx_files(lib_guide_dir)

    for f in lib_guide_files:
        print(f.name)

        # original version
        lib_guide_wb = load_workbook(Path(lib_guide_dir, f.name))
        lib_guide_ws = lib_guide_wb.worksheets[0]

        # final version (with faculty requests removed)
        pull_withdraw_wb = openpyxl.Workbook()
        pull_withdraw_ws = pull_withdraw_wb.active

        CALLNUMBER_COL_LIBGUIDE = None
        TITLE_COL_LIBGUIDE = None
        ISBN_COL_LIBGUIDE = None
        BARCODE_COL_LIBGUIDE = None
        WORLDCAT_COL_LIBGUIDE = None
        
        for row in lib_guide_ws.iter_rows(max_row = 1):
            # copy header row into updated withdrawal lists
            row_vals = [cell.value for cell in row]
            pull_withdraw_ws.append(row_vals)

            # locate the necessary columns they are not in the same location across spreadsheets
            for cell in row:
                col_header = cell.value
                if(col_header is not None):
                    col_header = col_header.lower()
                
                    if(col_header == 'display call number'):
                        CALLNUMBER_COL_LIBGUIDE = cell.column
                    elif(col_header == 'title'):
                        TITLE_COL_LIBGUIDE = cell.column
                    if(col_header == 'barcode'):
                        BARCODE_COL_LIBGUIDE = cell.column
                    elif(col_header == 'isbn'):
                        ISBN_COL_LIBGUIDE = cell.column
                    elif(col_header == 'worldcat oclc number'):
                        WORLDCAT_COL_LIBGUIDE = cell.column

        # move onto next file if there is no barcode data
        if(BARCODE_COL_LIBGUIDE is None):
            continue
                    
        # find the barcode in the faculty list
        for row in lib_guide_ws.iter_rows(min_row = 2):
            barcode = row[BARCODE_COL_LIBGUIDE - 1].value
            result = find_desired_val_from_search_val(barcode, faculty_keep_ws, BARCODE_COL_FACULTY, TRANSFERNAME_COL_FACULTY)
            
            # faculty has requested book, so add it to valid "pull & transfer" list
            if(result is not None):
                print(f'requester: {result} file: {f.name} barcode: {barcode}')

                # format specifically for the pull/transfer file
                row_vals = [row[CALLNUMBER_COL_LIBGUIDE - 1].value, row[TITLE_COL_LIBGUIDE - 1].value, row[ISBN_COL_LIBGUIDE - 1].value, row[BARCODE_COL_LIBGUIDE - 1].value, row[WORLDCAT_COL_LIBGUIDE - 1].value, result]

                # add row values to the "pull & transfer" spreadsheet
                pull_transfer_ws.append(row_vals)
                
            # faculty has not requested book (OR faculty did not format request properly), so add it to "pull & withdraw" list
            else:
                row_vals = [cell.value for cell in row]

                # add row to the "pull & withdraw" spreadsheet
                pull_withdraw_ws.append(row_vals)
                    
        # save every final "pull & withdraw" list 
        pull_withdraw_wb.save(filename = os.path.join(pull_withdraw_list_dir, f.name))

    # save final "pull & transfer" list
    pull_transfer_wb.save(filename = os.path.join(os.getcwd(), 'output', 'pull_transfer_lists', 'pull_transfer.xlsx'))

def log_not_found():
    BARCODE_COL_LIBGUIDE = column_index_from_string("BC")

    pull_transfer_file = Path(pull_withdraw_list_dir, 'pull_transfer.xlsx')
    pull_transfer_wb = load_workbook(pull_transfer_file)
    pull_transfer_ws = pull_transfer_wb.worksheets[0]

    not_found_log_wb = openpyxl.Workbook()
    not_found_log_ws = not_found_log_wb.active

    # copy header row into log
    for row in faculty_keep_ws.iter_rows(max_row = 1):
        row_vals = [cell.value for cell in row]
        not_found_log_ws.append(row_vals)

    for row in faculty_keep_ws.iter_rows(min_row = 2):
        barcode = row[BARCODE_COL_LIBGUIDE - 1].value
        result = find_desired_val_from_search_val(barcode, pull_transfer_ws, BARCODE_COL_FACULTY, TRANSFERNAME_COL_FACULTY)

        if(result is None):
            print(row[0].row, row[BARCODE_COL_FACULTY - 1].value, row[TRANSFERNAME_COL_FACULTY - 1].value)
            row_vals = [cell.value for cell in row]
            not_found_log_ws.append(row_vals)

    not_found_log_wb.save(filename = os.path.join(pull_withdraw_list_dir, "not_found_log2.xlsx"))


if __name__ == "__main__":
    start = time()
    main()
    end = time()
    print(f'[Processing time: {end - start} sec]')
