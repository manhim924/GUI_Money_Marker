import json 
import os
import config

TXT_FOLDER_PATH = config.path.INPUT_FOLDER
JSON_FILE_PATH = config.path.SERVER_OUTPUT_FILE

 
ALL_FILE_LIST = list(map(lambda files_: os.path.join(TXT_FOLDER_PATH, files_) ,config.path.EXCEL_INPUT_FILE_LIST))

def save_date(date):
    for file_path in ALL_FILE_LIST:
        with open(file_path ,'a',encoding='utf-8') as file:
            file.write(date +'\n')
                
                
def save_to_file(file_path, message):
    with open(file_path, 'a', encoding=('utf-8')) as file:
        file.write(message + ' \n')

def money_message(json, file):
    message = (
        f"{json.get("amount")} "
        f"\"{json.get("from")}\" "
        f"\"{json.get("to")}\" "
        f"{json.get("countable")} "
    )
    save_to_file(file, message)

def money_curreny_have(json, file):
    message =''
    for account , value in json.get("message").items():
        message  += f"{account} {value}\n"
    save_to_file(file,message)

def work_start(json, file):
    message = (
        "start "
        f"{json.get("time")} "
        f"\"{json.get("work for")}\""
    )
    save_to_file(file,message)

def work_end(json, file):
    message = (
        "end "
        f"{json.get("time")}"
        f"\"{json.get("work for")}\""
    )
    save_to_file(file,message)
    
def sleep_start(json, file):
    message = (
        "start "
        f"{json.get("time")}"
    )
    save_to_file(file,message)

def sleep_end(json, file):
    message = (
        "end "
        f"{json.get("time")}"
    )
    save_to_file(file,message)

def feeling(json, file):
    message = f"\"{json.get("feel")}\""
    save_to_file(file,message)

def water_fill(json, file):
    message = f"\"{json.get("container")}\""
    save_to_file(file,message)

def convert():
    with open(JSON_FILE_PATH,'r',encoding='utf-8') as json_file:
        data = json.load(json_file)

    for date_key , data_list in data.items():
        save_date(date_key)

        if(isinstance(data_list , list)):
            action_list = data_list
        else:
            action_list = [data_list]
        
        for each_action in action_list:
            operation = each_action.get("operation")

            if (operation == "Money message"):
                money_message(each_action, ALL_FILE_LIST[0])
            elif(operation == "Money current have"):
                money_curreny_have(each_action, ALL_FILE_LIST[1])                 
            elif(operation == "Work start"):
                work_start(each_action, ALL_FILE_LIST[2]) 
            elif(operation == "Work end"):
                work_end(each_action, ALL_FILE_LIST[2])
            elif(operation == "Sleep start"):
                sleep_start(each_action, ALL_FILE_LIST[3])
            elif(operation == "Sleep end"):
                sleep_end(each_action, ALL_FILE_LIST[3])
            elif(operation == "Feeling"):
                feeling(each_action, ALL_FILE_LIST[4]) 
            elif(operation == "Water fill"):
                water_fill(each_action, ALL_FILE_LIST[5])
    open(JSON_FILE_PATH,'w', encoding="utf-8").close()

if __name__ == "__main__":
    convert()