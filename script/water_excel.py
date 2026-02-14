from openpyxl import workbook
import config
import os

INPUT_FOLDER_PATH = config.path.INPUT_FOLDER
OUTPUT_FOLDER_PATH = config.path.OUTPUT_FOLDER

message_out = print
def set_message_out(function):
    global message_out
    message_out = function

def check_file_exist():
    pass

def read_data():
    # with open()
    pass

def water_excel_process():
    print("water excel started!")
    message_out("Water excel processing")
    pass

def main():
    check_file_exist()

if __name__ == "__main__":
    main()