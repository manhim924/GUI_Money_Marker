from flask import Flask, request, jsonify
import json
import config
import os

# --- verable setup start --- #
SAVE_FILE_PATH = config.path.SERVER_OUTPUT_FILE
PORT = config.server.SERVER_PORT_NUM
folder_path = os.path.dirname(SAVE_FILE_PATH)
if folder_path and not os.path.exists(folder_path):
    os.makedirs(folder_path)
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
server = Flask(__name__)

    # --- server accessable test start --- #
@server.route('/status', methods=['GET'])
def status_check():
    return jsonify({"status":"online"}), 200
    # --- server accessable test end --- #

    # --- receive file function start --- #
@server.route('/upload', methods=['POST'])
def get_file():
    with open(SAVE_FILE_PATH, "w", encoding="utf-8") as f:
        pass
    try:
        data = request.get_json(force=True) # get file
        
        if not data: 
            data = json.loads(request.data.decode('utf-8'))  
        
        if not data:
            return jsonify({"status": "error", "message":"No data found"}), 400

        file_old_data = []
        if os.path.exists(SAVE_FILE_PATH):
            try:
                with open(SAVE_FILE_PATH, 'r', encoding='utf-8') as file:
                    file_old_data = json.load(file)
            except json.JSONDecodeError:
                file_old_data = []

        if isinstance(data, list):
            file_old_data.extend(data)
        else:
            file_old_data.append(data)
        
        with open(SAVE_FILE_PATH, 'w', encoding='utf-8') as file:
            json.dump(file_old_data, file, indent=4 , ensure_ascii=False)

        message_output(f"Server : Rreceived data! Saved to {os.path.basename(SAVE_FILE_PATH)}")

        if json_converter:
            json_converter()
            message_output(f"Server : Json convereted")

        return jsonify({"status": "success"}), 200
        
    except Exception as error:
        message_output(f"Server Error: {str(error)}")
        return jsonify({"status": "error", "message": str(error)}), 500
    # --- receive file function start --- #
    
def server_start():
    
    print("Old json file is cleared")
    print("The server is started!")
    server.run(host='0.0.0.0', port=PORT)

if __name__ == "__main__":
    # import json_converter
    # set_json_converter(json_converter.convert)
    server_start()