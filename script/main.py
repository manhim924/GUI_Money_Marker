import window_GUI
import server
import feeling_excel
import money_excel
import sleep_excel
import water_excel
import work_excel
import json_converter
from sys import argv as sys_argv
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QThread, pyqtSignal

# --- parallel for pyqt6 start --- #
class Server_parallel(QThread):
  log_signal = pyqtSignal(str)

  def send_log_to_gui(self, message):
    self.log_signal.emit(message)

  def run(self):
    self.log_signal.emit("Sysyem: Attempting to start server...")
    server.set_message_output(self.send_log_to_gui)
    try:
      server.set_json_converter(json_converter.convert)
      server.server_start()
    except Exception as error:
      self.log_signal.emit(f"Server error: {str(error)}")

  def stop_server(self):
    server.server_stop()
    self.quit()
    self.wait()
# --- parallel for pyqt6 end  --- #

def main():
  app = QApplication(sys_argv)

  window = window_GUI.Main_window(1920,1080)
  parallel = Server_parallel()

  message_output = window.log_window_message_output

  window.server_start_signal.connect(parallel.start)
  window.server_stop_signal.connect(parallel.stop_server)

  parallel.log_signal.connect(message_output)
  
  excel_modules = [
    (money_excel, window.money_excel_start_signal, money_excel.money_excel_process),
    (feeling_excel, window.feeling_excel_start_signal, feeling_excel.feeling_excel_process),
    (sleep_excel, window.sleep_excel_start_signal, sleep_excel.sleep_excel_process),
    (water_excel, window.water_excel_start_signal, water_excel.water_excel_process),
    (work_excel, window.work_excel_start_signal, work_excel.work_excel_process),
  ]

  for module, signal, function in excel_modules:
    signal.connect(function)
    module.set_message_out(message_output)

  window.show()
  app.exec()

if __name__ == '__main__':
  main()