from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl import Workbook, load_workbook
from datetime import datetime
from os import (path as os_path)

class Color_style():
    PALE_MINT = PatternFill(fill_type="solid", fgColor="EBF1DE") 
    PALE_PINK = PatternFill(fill_type="solid", fgColor="F2DCDB")
    LIGHT_BLUE = PatternFill(fill_type="solid",fgColor="DCE6F1")
    SKY_BLUE = PatternFill(fill_type="solid",fgColor="C5D9F1")
    PEACH = PatternFill(fill_type="solid", fgColor="FDE9D9") 
    LACENDER = PatternFill(fill_type="solid", fgColor="E4DFEC")
    BEIGE = PatternFill(fill_type="solid",fgColor="DDD9C4")
    LIGHT_CORAL = PatternFill(fill_type="solid",fgColor="FFBDBD")
    YELLOW_GREEN = PatternFill(fill_type="solid",fgColor="92D050")
    SATIN_GREEN = PatternFill(fill_type="solid", fgColor="C6E0B4") 
    SPEARMINT_GREEN = PatternFill(fill_type="solid", fgColor="19ff78")
    FLIRTY_SALMON = PatternFill(fill_type="solid", fgColor="ff6d6d")

class Font_style():
    FONT = Font(name='Book Antiqua', size=12)

class Border_style():
    S_M = Side(border_style="medium")
    S_D = Side(border_style="dashed")
    S_T = Side(border_style="thin")

    border_dict = {
        "M" : S_M,
        "D" : S_D,
        "T" : S_T
    }

    @classmethod
    def set(cls,t=None,r=None,l=None,b=None): 
        return Border(
            top = cls.border_dict.get(t),         
            bottom = cls.border_dict.get(b),         
            left = cls.border_dict.get(l),         
            right = cls.border_dict.get(r)         
        )

class Function():
    def read_first_line_of_file(self,file):
        with open(file, 'r', encoding='utf-8') as f:
            return f.readline().strip()

    def is_date(self, data):
        try:
            datetime.strptime(data, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def check_and_open_excel(self, file):
        if os_path.exists(file):
            wb = load_workbook(file)
        else:
            wb = Workbook()
        return wb

    def read_file_data(self, file):
        with open(file,'r', encoding="utf-8") as file:
            data = file.read().splitlines()
        return data

    def month_int_to_string(self, month):
        month_dict = {
            1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
            7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
        }
        return month_dict.get(month)

    def open_month_ws(self,wb,ws):
        if ws in wb.sheetnames:
            return wb[ws]
        else:
            return wb.create_sheet(ws)

    def get_max_row_by_value(self, ws):
        for row in range(ws.max_row, 0, -1):
            if any(cell.value is not None for cell in ws[row]):
                return row
        return 0

    def is_first_time_of_ws(self,ws, pass_number):
        last_row = self.get_max_row_by_value(ws)
        for i in range(last_row,1,-1):
            cell = f"B{i}"
            value = ws[cell].value
            if value != None:
                if self.is_date(ws[cell].value):
                    if(pass_number == 0):
                        return False
                    else:
                        pass_number-=1
        return True