import os
from datetime import datetime


class path():
    CURRENT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
    PROJECT_ROOT_FOLDER =os.path.dirname(CURRENT_DIRECTORY)

    INPUT_FOLDER = os.path.join(PROJECT_ROOT_FOLDER, "input")
    OUTPUT_FOLDER = os.path.join(PROJECT_ROOT_FOLDER, "output")

    SERVER_OUTPUT_FILE = os.path.join(PROJECT_ROOT_FOLDER,"data_received", "server_output.json") 

    GUI_SETTING_FILE = os.path.join(PROJECT_ROOT_FOLDER, "config", "setting.txt")

    EXCEL_INPUT_FILE_LIST = ["money_messages.txt","money_current_have.txt","work_message.txt","sleep_message.txt","feeling_message.txt","water_message.txt"]

class server():
    SERVER_PORT_NUM = 5000

def test():
    print(path.PROJECT_ROOT_FOLDER)

if __name__ == "__main__":
    test()