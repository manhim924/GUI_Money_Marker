from PyQt6.QtWidgets import (QApplication,  QWidget , QMainWindow, QPushButton, 
                             QVBoxLayout, QHBoxLayout, QSplitter, QStackedWidget, 
                             QTextEdit, QLabel)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction
from datetime import datetime
from platform import system as platform_system
from os import (path as os_path,
                makedirs as os_makedirs,
                remove as os_remove,
                rename as os_rename)

if( platform_system() == "Windows"):
    from os import  startfile as os_startfile

from sys import argv as sys_argv
from json import (dump as json_dump,
                  load as json_load)
from subprocess import run as subprocess_run
import config 

class Main_window(QMainWindow):
    BUTTON_STYLE_SHEET=("""
        QPushButton{
            color: #ffffff;
            background-color: #55000000;
            font-size: 20px;
            border: 1px solid #000000;
            border-radius : 5px;
        }

        QPushButton:hover{
            background-color: #AA000000;
        }
    """)
    LOG_WINDOW_STYLE_SHEET=(""" 
        QTextEdit{
            background-color : #002D07;
            color: #00ff00;
            font-family: Consolas, Monospace;
            font-size: 20px;
            border: 3px solid #0f0f0f;                          
            margin: 10px;
            border-radius : 5px;
        } 
    """)
    TEXT_STYLE_SHEET=("""
        QLabel{
           font-size: 20px;               
           font-weight: 700;
        }
    """)
    USER_STYLE_SHEET=("""
        QTextEdit{
            font-size: 20px;
            border:3px solid #0f0f0f;
            font-family: Georgia;
            margin: 10px;
            border-radius : 5px;
        }
    """)

    CURRENT_DIRECTORY = config.path.PROJECT_ROOT_FOLDER
    SETTING_FILE =  config.path.GUI_SETTING_FILE
    INPUT_FOLDER = config.path.INPUT_FOLDER
    OUTPUT_FOLDER = config.path.OUTPUT_FOLDER

    # --- signal for main.py setup start --- #
    current_opening_file = None
    server_start_signal = pyqtSignal()
    money_excel_start_signal = pyqtSignal()
    work_excel_start_signal = pyqtSignal()
    sleep_excel_start_signal = pyqtSignal()
    water_excel_start_signal = pyqtSignal()
    feeling_excel_start_signal = pyqtSignal()
    # --- singal for main.py setup end --- #        

    def __init__(self, width, height):
        super().__init__()
        # --- basic window setting start --- # 
        self.width = width
        self.height = height
        self.setWindowTitle("GUI Daily Marker")
        self.resize(width, height) 
        # --- basic window setting end --- #

        # --- menu bar start --- #
        menu_bar = self.menuBar()
        view_button = menu_bar.addMenu("&View")

        view_action_1 = QAction("Show money marking", self)
        view_action_1.setShortcut("Ctrl+1")
        view_action_1.triggered.connect(lambda :self.l_stacked_w.setCurrentIndex(0))
        view_button.addAction(view_action_1)

        view_action_2 = QAction("Show daily marking", self)
        view_action_2.setShortcut("Ctrl+2")
        view_action_2.triggered.connect(lambda :self.l_stacked_w.setCurrentIndex(1))
        view_button.addAction(view_action_2)

        view_action_3 = QAction("Show all excel reload", self)
        view_action_3.setShortcut("Ctrl+3")
        view_action_3.triggered.connect(lambda :self.l_stacked_w.setCurrentIndex(2))
        view_button.addAction(view_action_3)
        
        view_action_4 = QAction("Show all excel reload", self)
        view_action_4.setShortcut("Ctrl+4")
        view_action_4.triggered.connect(lambda :self.l_stacked_w.setCurrentIndex(3))
        view_button.addAction(view_action_4)
        # --- menu bar end --- #

        # --- left main widget money marking button start --- #
        server_start_button = self.Button("Server Start", True, [self.server_start_signal.emit])
        server_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        money_excel_start_button = self.Button("Money Excel Start", True, [self.money_excel_start_signal.emit]) 
        money_excel_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET) 

        load_current_have = self.Button("Load Current Have", True,  [lambda: self.Load_file("money_current_have.txt")])
        load_current_have.setStyleSheet(self.BUTTON_STYLE_SHEET)

        load_money_message = self.Button("Load Money Message", True, [lambda: self.Load_file("money_messages.txt")])
        load_money_message.setStyleSheet(self.BUTTON_STYLE_SHEET)

        l_money_marking_layout= QVBoxLayout() 
        l_money_marking_layout.addWidget(server_start_button)
        l_money_marking_layout.addWidget(money_excel_start_button)
        l_money_marking_layout.addWidget(load_money_message)
        l_money_marking_layout.addWidget(load_current_have)

        l_money_marking_widget_set = self.Color_widget(layout=l_money_marking_layout, bgc="#6d6d6d", border="2px solid #000000")  
        # --- left main widget money marking button end --- #

        # --- left main widget daily marking button start --- #
        water_start_button = self.Button("Water Start Mark", True,  [self.water_excel_start_signal.emit])
        water_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        feeling_start_button = self.Button("Feeling Start Mark", True,  [self.feeling_excel_start_signal.emit])
        feeling_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        working_start_button = self.Button("Working Start Mark", True, [self.work_excel_start_signal.emit])
        working_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        sleeping_start_button = self.Button("Sleeping Start Mark", True, [self.sleep_excel_start_signal.emit])
        sleeping_start_button.setStyleSheet(self.BUTTON_STYLE_SHEET)
        
        l_daily_marking_layout = QVBoxLayout()
        l_daily_marking_layout.addWidget(water_start_button)
        l_daily_marking_layout.addWidget(feeling_start_button)
        l_daily_marking_layout.addWidget(working_start_button)
        l_daily_marking_layout.addWidget(sleeping_start_button)

        l_daily_marking_widget_set = self.Color_widget(layout=l_daily_marking_layout, bgc="#6d6d6d", border="2px solid #000000")  
        # --- left main widget daily marking button end --- #

        # --- left main widget all reload button start --- #
        current_year = datetime.now().year
        current_year_file_path = os_path.join(self.OUTPUT_FOLDER, str(current_year))

        money_excel_reload_button = self.Button("Reload Money Excel", True, [lambda: self.reload_file(os_path.join(current_year_file_path,"money"), "money.xlsx"), lambda: self.log_window_message_output("Money excel reloaded")])
        money_excel_reload_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        work_excel_reload_button = self.Button("Reload Work Excel", True, [lambda: self.reload_file(os_path.join(current_year_file_path,"work"), "work.xlsx"), lambda: self.log_window_message_output("Work excel reloaded")])
        work_excel_reload_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        sleep_excel_reload_button = self.Button("Reload Sleep Excel", True, [lambda: self.reload_file(os_path.join(current_year_file_path, "sleep"),"sleep.xlsx"), lambda: self.log_window_message_output("Sleep excel reloaded")])
        sleep_excel_reload_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        feeling_excel_reload_button = self.Button("Reload Feeling Excel", True, [lambda: self.reload_file(os_path.join(current_year_file_path,"feeling"),"feeling.xlsx"), lambda: self.log_window_message_output("Feeling excel reloaded")])
        feeling_excel_reload_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        water_excel_reload_button = self.Button("Reload Water Excel", True, [lambda: self.reload_file(os_path.join(current_year_file_path,"water"),"water.xlsx"), lambda: self.log_window_message_output("Water excel reloaded")])
        water_excel_reload_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        l_reload_excel_layout = QVBoxLayout()
        l_reload_excel_layout.addWidget(money_excel_reload_button)
        l_reload_excel_layout.addWidget(work_excel_reload_button)
        l_reload_excel_layout.addWidget(sleep_excel_reload_button)
        l_reload_excel_layout.addWidget(feeling_excel_reload_button)
        l_reload_excel_layout.addWidget(water_excel_reload_button)

        l_reload_excel_widget_set = self.Color_widget(layout=l_reload_excel_layout, bgc="#6d6d6d", border="2px solid #000000")
        # --- left main widget all reload button end  --- #

        # --- left main widget open file butten start --- #
        money_excel_open_button = self.Button("Open Money Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "money", "money.xlsx"))])
        money_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        money_backup_excel_open_button = self.Button("Open Backup Money Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "money", "backup.xlsx"))])
        money_backup_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        work_excel_open_button = self.Button("Open Work Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "work", "work.xlsx"))])
        work_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        sleep_excel_open_button = self.Button("Open Sleep Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "sleep", "sleep.xlsx"))])
        sleep_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        feeling_excel_open_button = self.Button("Open Feeling Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "feeling", "feeling.xlsx"))])
        feeling_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        water_excel_open_button = self.Button("Open Water Excel", True, [lambda: self.open_file(os_path.join(current_year_file_path, "water", "water.xlsx"))])
        water_excel_open_button.setStyleSheet(self.BUTTON_STYLE_SHEET)
        
        l_open_excel_layout = QVBoxLayout()
        l_open_excel_layout.addWidget(money_excel_open_button)
        l_open_excel_layout.addWidget(money_backup_excel_open_button)
        l_open_excel_layout.addWidget(work_excel_open_button)
        l_open_excel_layout.addWidget(sleep_excel_open_button)
        l_open_excel_layout.addWidget(feeling_excel_open_button)
        l_open_excel_layout.addWidget(water_excel_open_button)

        l_open_excel_widget_set = self.Color_widget(layout=l_open_excel_layout, bgc="#6d6d6d", border="2px solid #000000")

        # --- left main widget open file butten end --- #

        # --- left stacked widget start --- #
        self.l_stacked_w = QStackedWidget()
        self.l_stacked_w.addWidget(l_money_marking_widget_set) # index 0
        self.l_stacked_w.addWidget(l_daily_marking_widget_set) # index 1
        self.l_stacked_w.addWidget(l_reload_excel_widget_set)  # index 2
        self.l_stacked_w.addWidget(l_open_excel_widget_set) # index 3
        # --- left stacked widget end --- #

        # --- log widget start--- #
        log_window_text = QLabel("Application Message")
        log_window_text.setStyleSheet(self.TEXT_STYLE_SHEET)
        log_window_text.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        self.log_window.setStyleSheet(self.LOG_WINDOW_STYLE_SHEET)

        right_top_layout = QVBoxLayout()
        right_top_layout.addWidget(log_window_text)
        right_top_layout.addWidget(self.log_window)

        right_top_widget = self.Color_widget(layout = right_top_layout, bgc="#b5b5b5", border="2px solid #000000", margin="5px") 
        # --- log widget end --- # 

        # --- folder checking start --- #

        for folder in ["config", "input", "output"]:
            folder_path = os_path.join(self.CURRENT_DIRECTORY, folder)
            if not os_path.exists(folder_path):
                self.log_window_message_output(f"System: Folder \"{folder}\" not found")
                os_makedirs(folder_path)
                self.log_window_message_output(f"System: Folder \"{folder}\" is created") 

        # --- folder checking end --- #

        # --- user input widget start --- #    
        user_input_text = QLabel("User Input")
        user_input_text.setStyleSheet(self.TEXT_STYLE_SHEET)
        user_input_text.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        self.user_input_window = QTextEdit()
        self.user_input_window.setStyleSheet(self.USER_STYLE_SHEET)

        user_input_save_button = self.Button("Save",True, [lambda: self.Save_file()]) 
        user_input_save_button.setShortcut("Ctrl+s")
        user_input_save_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        clear_user_input_window_button = self.Button("Clear", True, [lambda: self.user_input_window.clear()] )
        clear_user_input_window_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        start_all_job_button = self.Button("Start All Excel", True, [
            self.money_excel_start_signal.emit,
            self.work_excel_start_signal.emit,
            self.sleep_excel_start_signal.emit,
            self.water_excel_start_signal.emit,
            self.feeling_excel_start_signal.emit            
        ])
        start_all_job_button.setStyleSheet(self.BUTTON_STYLE_SHEET)

        right_bottom_button_layout = QHBoxLayout()
        right_bottom_button_layout.addWidget(user_input_save_button)
        right_bottom_button_layout.addWidget(clear_user_input_window_button)
        right_bottom_button_layout.addWidget(start_all_job_button)

        right_bottom_layout= QVBoxLayout()
        right_bottom_layout.addWidget(user_input_text)
        right_bottom_layout.addWidget(self.user_input_window)
        right_bottom_layout.addLayout(right_bottom_button_layout)

        right_bottom_widget = self.Color_widget(layout=right_bottom_layout, bgc="#b5b5b5", border="2px solid #000000",margin="5px") 

        # --- user input widget end --- #
        self.right_main_splitter = QSplitter(Qt.Orientation.Vertical)
        self.right_main_splitter.addWidget(right_top_widget)
        self.right_main_splitter.addWidget(right_bottom_widget)
        self.right_main_splitter.setSizes([540,540])

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.addWidget(self.l_stacked_w)
        self.main_splitter.addWidget(self.right_main_splitter)

        self.main_splitter.setSizes([384,1536])
        self.setCentralWidget(self.main_splitter)

        self.Load_setting()

    def closeEvent(self, event):
        self.Save_settings()
        super().closeEvent(event)

    def Save_settings(self):
        
        location = self.geometry()

        setting = {
            "x": location.x(),
            "y": location.y(),
            "width": location.width(),
            "height": location.height(),
            "main_splitter_sizes": self.main_splitter.sizes(),
            "right_main_splitter_sizes" : self.right_main_splitter.sizes()             
        }

        with open(self.SETTING_FILE,'w') as file:
            json_dump(setting, file, indent=4)
        print("Setting saved")

    def Load_setting(self):
        if not os_path.exists(self.SETTING_FILE):
            return
        
        with open(self.SETTING_FILE, 'r') as file:  
            data = json_load(file)

        if "x" in data and "y" in data:
            self.move(data["x"], data["y"])
        
        if "width" in data and "height" in data:
            self.resize(data["width"], data["height"])
        
        if "main_splitter_sizes" in data:
            self.main_splitter.setSizes(data["main_splitter_sizes"])

        if "right_main_splitter_sizes" in data:
            self.right_main_splitter.setSizes(data["right_main_splitter_sizes"])

    def Load_file(self, file_name):
        self.user_input_window.clear()
         
        file_path= os_path.join(self.INPUT_FOLDER, file_name)

        try:
            with open(file_path, 'r', encoding="utf-8") as file:
                data = file.read()
            self.user_input_window.append(data)
            self.log_window_message_output(f"File {file_name} is loaded!")

        except FileNotFoundError:
            self.log_window_message_output(f"File {file_name} not exists")

            with open(file_path, 'w',encoding="utf-8" ) as file:
                pass
            self.log_window_message_output(f"File {file_name} created")    

        self.current_opening_file = file_path
        self.user_input_window.document().clearUndoRedoStacks()

    def Save_file(self):
        try:
            text_to_save =  self.user_input_window.toPlainText()
            with open(self.current_opening_file, 'w',encoding="utf-8") as file:
                file.write(text_to_save)
            self.user_input_window.clear()
            self.log_window_message_output(f"File {os_path.basename(self.current_opening_file)} is saved!")
            self.current_opening_file = None
        except TypeError:
            self.log_window_message_output(f"Nothing is not opening!")
            self.log_window_message_output(f"Please open some file before saving!")      


    def Color_widget(self, layout ,border="",bgc = "#ffffff", margin=""):
        widget = QWidget()
        widget.setObjectName("this_widget")
        widget.setStyleSheet(f"""
            #this_widget{{
                background-color: {bgc};
                border : {border};
                margin : {margin};
            }}
        """)
        widget.setLayout(layout) 

        return widget

    def Button(self, name, can_click, functions):
        button = QPushButton(name)
        button.setCheckable(can_click)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        for func in functions:
            button.pressed.connect(func)
        return button

    def log_window_message_output(self, input, new_line = False):
        time = datetime.now().strftime("%H:%M:%S") 
        message = f"{time}: {input}"
        if(new_line):
            message += "\n"
        self.log_window.append(message) 

    def reload_file(self, file_path, file_name):
        remove_file_path = os_path.join(file_path, file_name)
        reload_file_path = os_path.join(file_path, "backup.xlsx")
        if(os_path.exists(remove_file_path) and os_path.exists(reload_file_path)):
            os_remove(remove_file_path)
            os_rename(reload_file_path, remove_file_path)
        else:
            self.log_window_message_output(f"file path \"{file_path}\" not exist")


    def open_file(self, file_path):
        if (not os_path.exists(file_path)):
            self.log_window_message_output(f"File {file_path} not exist")
            return
        
        current_os = platform_system()

        if (current_os == "Windows"):
            os_startfile(file_path)
        elif(current_os == "Darwin"): # MAC os
            subprocess_run(["open", file_path] )
        elif(current_os == "Linux"):
            subprocess_run(["xdg-open", file_path])
        else:
            self.log_window_message_output(f"Unsupproted OS : {current_os}") 

def test():
    app = QApplication(sys_argv)
    window = Main_window(1920,1080) 
    window.show()

    app.exec() 

if __name__ == "__main__":
    test()