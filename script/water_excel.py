from openpyxl import workbook
from os import (path as os_path)

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

def water_excel_process():
    message_out("Water excel started!")

    file_first_line = eh_funct.read_first_line_of_file(WATER_MESSAGE_FILE_PATH)
    if(eh_funct.is_date(file_first_line)):

        pass


if __name__ == "__main__":
    water_excel_process();