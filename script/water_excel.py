from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from sys import (exit as sys_exit)
from os import (path as os_path,
                makedirs as os_makedirs)
from datetime import datetime
from decimal import Decimal
from shutil import (copy as shutil_copy)
from shlex import (split as shlex_split)
import config
from excel_helper import (Color_style as eh_color,
                          Font_style as eh_font,
                          Border_style as eh_border,
                          Function as eh_funct)

INPUT_FOLDER_PATH = config.path.INPUT_FOLDER
WATER_MESSAGE_FILE_PATH = os_path.join(INPUT_FOLDER_PATH, config.path.EXCEL_INPUT_FILE_LIST[5])

OUTPUT_FOLDER_PATH = config.path.OUTPUT_FOLDER

message_out = print
def set_message_out(function):
    global message_out
    message_out = function

def sheet_init(wb, ws, month, year):
    ws.column_dimensions['B'].width = 14.5

    for i in range(3,23):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = 12

    cell_message = []

    cell_color = []

    cell_border = []

    merge_cell = []

    eh_funct.cell_type_message(ws, cell_message)
    eh_funct.range_cell_set_color(ws, cell_color)
    eh_funct.cell_set_border(ws, cell_border)
    eh_funct.range_merage_cell(ws, merge_cell)    

def water_excel_process():
    message_out("Water excel started!")

    file_first_line = eh_funct.read_first_line_of_file(WATER_MESSAGE_FILE_PATH)
    if(eh_funct.is_date(file_first_line)):
        f_day, f_month, f_year = file_first_line.split('/')
        WATER_FOLDER_PATH = os_path.join(OUTPUT_FOLDER_PATH, f_year, "water")
        WATER_FILE_PATH = os_path.join(WATER_FOLDER_PATH, "water.xlsx")
        BACKUP_FILE_PATH = os_path.join(WATER_FOLDER_PATH, "backup.xlsx")
    else:
        if(message_out):
            message_out("Water Excel : Error, the water_message.txt first line is not a date!")
            sys_exit(0)

    if(not os_path.exists(WATER_FOLDER_PATH)):
        os_makedirs(WATER_FOLDER_PATH)
    
    if(os_path.exists(WATER_FILE_PATH)):
        shutil_copy(WATER_FILE_PATH, BACKUP_FILE_PATH)

    wb = eh_funct.check_and_open_excel(WATER_FILE_PATH)
    ws = wb.active

    current_date = None
    last_day = None
    water_message_list = eh_funct.read_file_data(WATER_MESSAGE_FILE_PATH)
    START_COL = 2
    have_previous_date = False
    can_save = True
    for message in water_excel_process:
        if(message == ''):
            continue

        if eh_funct.is_date(message):
            if(current_date != None):
                last_day = current_date

            if(have_previous_date):
                # day_finish(ws, last_day, start_row + record)
                pass

            day, month, year = message.split('/')
            month_str = eh_funct.month_int_to_string(int(month))
            ws = eh_funct.open_month_ws(wb, month_str)

            current_date = message

            if(eh_funct.is_first_time_of_ws(ws,0)):
                sheet_init(wb, ws, month, year)
                pass



if __name__ == "__main__":
    water_excel_process();