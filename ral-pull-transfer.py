from fileinput import filename
import os
from time import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
import openpyxl
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from isbncommon import *

CALLNUMBER_COL_LIBGUIDE = column_index_from_string("B")
BARCODE_COL_LIBGUIDE = column_index_from_string("BC")
ISBN_COL_LIBGUIDE = column_index_from_string("BB")

CALLNUMBER_COL_FACULTY = column_index_from_string("B")
BARCODE_COL_FACULTY = column_index_from_string("BC")
ISBN_COL_FACULTY = column_index_from_string("BB")
TRANSFERNAME_COL_FACULTY = column_index_from_string("BE")


def main():
    lib_guide_dir = Path(os.getcwd(), "input", "lib_guide_title_lists")
    faculty_keep_dir = Path(
        os.getcwd(), "input", "faculty_keep_lists", "faculty_keep_masterlist.xlsx"
    )
    pull_list_dir = Path(os.getcwd(), "output", "pull_lists")

    # master list of all faculty requests compiled via email
    faculty_keep_wb = load_workbook(Path(faculty_keep_dir))
    faculty_keep_ws = faculty_keep_wb.worksheets[0]

    # master list of all faculty requests found in the withdrawal lists that will be transfered to their offices
    pull_transfer_wb = openpyxl.Workbook()
    pull_transfer_ws = pull_transfer_wb.active

    # copy header row into "pull & transfer" master list
    for row in faculty_keep_ws.iter_rows(max_row = 1):
        row_vals = [cell.value for cell in row]
        pull_transfer_ws.append(row_vals)

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

        # copy header row into updated withdrawal lists
        for row in lib_guide_ws.iter_rows(max_row = 1):
            row_vals = [cell.value for cell in row]
            pull_withdraw_ws.append(row_vals)

        # find the call number in the faculty list
        for row in lib_guide_ws.iter_rows(min_row = 2):
            call_number = row[CALLNUMBER_COL_LIBGUIDE - 1].value
            result = find_desired_val_from_search_val(call_number, faculty_keep_ws, CALLNUMBER_COL_FACULTY, TRANSFERNAME_COL_FACULTY)
            
            # faculty has requested book, so add it to valid "pull & transfer" list
            if(result is not None):
                print(f'requester: {result} file: {f.name} call number: {call_number}')
                row_vals = [cell.value for cell in row]
                
                # add requester name to the data set
                row_vals.append(result)

                # add row to the "pull & transfer" spreadsheet
                pull_transfer_ws.append(row_vals)
                
            # faculty has not requested book (OR faculty did not format request properly), so add it to "pull & withdraw" list
            else:
                row_vals = [cell.value for cell in row]

                # add row to the "pull & withdraw" spreadsheet
                pull_withdraw_ws.append(row_vals)
        
        # save every final "pull & withdraw" list 
        pull_withdraw_wb.save(filename = os.path.join(pull_list_dir, f.name))

    # save final "pull & transfer" list
    pull_transfer_wb.save(filename = os.path.join(pull_list_dir, "pull_transfer.xlsx"))


if __name__ == "__main__":
    main()
