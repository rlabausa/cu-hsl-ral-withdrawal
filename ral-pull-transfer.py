from fileinput import filename
import os
from time import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
import openpyxl
from openpyxl.utils.cell import column_index_from_string       

from isbncommon import *

CALLNUMBER_LIBGUIDE = column_index_from_string('B')
BARCODE_COL_LIBGUIDE = column_index_from_string('BC')
ISBN_COL_LIBGUIDE = column_index_from_string('BB')

CALLNUMBER_FACULTY = column_index_from_string('B')
BARCODE_COL_FACULTY = column_index_from_string('BC')
ISBN_COL_FACULTY = column_index_from_string('BB')
TRANSFERNAME_FACULTY = column_index_from_string('BE')

def main():
    lib_guide_dir = Path(os.getcwd(),'input', 'lib_guide_title_lists')
    faculty_keep_dir = Path(os.getcwd(), 'input', 'faculty_keep_lists', 'faculty_keep_masterlist.xlsx')

    faculty_keep_wb = load_workbook(Path(faculty_keep_dir))
    faculty_keep_ws = faculty_keep_wb.worksheets[0]

    lib_guide_files = get_xlsx_files(lib_guide_dir)

    for f in lib_guide_files:
        print(f.name)

        new_lib_guide_wb = openpyxl.Workbook()
        new_lib_guide_ws = new_lib_guide_wb.active

        lib_guide_wb = load_workbook(Path(lib_guide_dir, f.name))
        lib_guide_ws = lib_guide_wb.worksheets[0]

        # TODO: Find the call#/isbn/worldcat num in the faculty list and omit in new worksheet if found
        for row in lib_guide_ws.iter_rows(min_row=2):
            for cell in row:
                print (cell.value)
                
        fout = f'output/pull_lists/{f.name}'
        new_lib_guide_wb.save(filename=fout)
        

   


if (__name__ == '__main__'):
    main();
