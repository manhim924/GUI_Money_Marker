from flask import (Flask as flask_Flask,
                   request as flask_request,
                   jsonify as flask_jsonify)
from json import (loads as json_loads,
                  load as json_load,
                  dump as json_dump,
                  JSONDecodeError as json_JSONDecodeError)
from os import (path as os_path,
                makedirs as os_makedirs)
from shutil import (copy2 as shutil_copy2)
from werkzeug.serving import make_server

import config

# --- verable setup start --- #
SAVE_FILE_PATH = config.path.SERVER_OUTPUT_FILE
PORT = config.server.SERVER_PORT_NUM
folder_path = os_path.dirname(SAVE_FILE_PATH)
if folder_path and not os_path.exists(folder_path):
    os_makedirs(folder_path)
flask_server_instance = None
# --- verable setup end --- #

# --- setup function for main.py start --- #
message_output = print
def set_message_output(function):
    global message_output
    message_output = function

json_converter = None
def set_json_converter(function):
    global json_converter
    json_converter = function
# --- setup function for main.py end --- #

# --- main function start --- # 
server = flask_Flask(__name__)

    # --- server accessable test start --- #
@server.route('/status', methods=['GET'])
def status_check():
    return flask_jsonify({"status":"online"}), 200
    # --- server accessable test end --- #

    # --- receive file function start --- #
@server.route('/upload', methods=['POST'])
def get_file():
    if os_path.exists(SAVE_FILE_PATH):
        backup_file_path = os_path.join(os_path.dirname(SAVE_FILE_PATH), "backup.json")
        shutil_copy2(SAVE_FILE_PATH, backup_file_path)
    else:
        print(f"Skipping backup: {SAVE_FILE_PATH} not found.")

    message_output("Server file backuped") 

    with open(SAVE_FILE_PATH, "w", encoding="utf-8") as f:
        pass

    message_output("Server file clear")

    try:
        data = flask_request.get_json(force=True) # get file
        
        if not data: 
            data = json_loads(flask_request.data.decode('utf-8'))  
        
        if not data:
            return flask_jsonify({"status": "error", "message":"No data found"}), 400

        file_old_data = []
        if os_path.exists(SAVE_FILE_PATH):
            try:
                with open(SAVE_FILE_PATH, 'r', encoding='utf-8') as file:
                    file_old_data = json_load(file)
            except json_JSONDecodeError:
                file_old_data = []

        if isinstance(data, list):
            file_old_data.extend(data)
        else:
            file_old_data.append(data)
        
        with open(SAVE_FILE_PATH, 'w', encoding='utf-8') as file:
            json_dump(file_old_data, file, indent=4 , ensure_ascii=False)

        message_output(f"Server : Rreceived data! Saved to {os_path.basename(SAVE_FILE_PATH)}")

        if json_converter:
            json_converter()
            message_output(f"Server : Json convereted")

        return flask_jsonify({"status": "success"}), 200
        
    except Exception as error:
        message_output(f"Server Error: {str(error)}")
        return flask_jsonify({"status": "error", "message": str(error)}), 500
    # --- receive file function start --- #
    
def server_start():
    global flask_server_instance
    message_output("Server: server started!")

    flask_server_instance = make_server("0.0.0.0", PORT, server)

    flask_server_instance.serve_forever()

def server_stop():
    global flask_server_instance
    if flask_server_instance:
        flask_server_instance.shutdown()
        message_output("Server: server stoped!")
    else:
        message_output("Server: server not started, cannot stop", color="red")


if __name__ == "__main__":
    # import json_converter
    # set_json_converter(json_converter.convert)
    server_start()