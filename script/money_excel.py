from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import sys
import os
from datetime import datetime, timedelta
from decimal import Decimal
import shlex
import config
import excel_helper as helper

INPUT_FOLDER_PATH = config.path.INPUT_FOLDER
MONEY_MESSAGE_FILE_PATH = os.path.join(INPUT_FOLDER_PATH, config.path.EXCEL_INPUT_FILE_LIST[0])  
CURRENT_HAVE_FILE_PATH = os.path.join(INPUT_FOLDER_PATH, config.path.EXCEL_INPUT_FILE_LIST[1])

OUTPUT_FOLDER_PATH = config.path.OUTPUT_FOLDER


message_out = print
def set_message_out(function):
    global message_out
    message_out = function

def read_first_line_of_file(file):
    with open(file, 'r', encoding='utf-8') as f:
        return f.readline().strip()

def is_folder_exist(folder):
    if not os.path.exists(folder):
        return False
    return True

def check_and_open_excel(file):
    if is_folder_exist(file):
        wb = load_workbook(file)
    else:
        wb = Workbook()
    return wb

def read_file_data(file):
    with open(file,'r', encoding="utf-8") as file:
        data = file.read().splitlines()
    return data

def is_date(data):
    try:
        datetime.strptime(data, "%d/%m/%Y")
        return True
    except ValueError:
        return False

def month_to_string(month):
    month_dict = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
        "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
    }
    return month_dict.get(month)

def open_month_ws(wb,ws):
    if ws in wb.sheetnames:
        return wb[ws]
    else:
        return wb.create_sheet(ws)


def is_first_time_of_ws(ws):
    last_row = ws.max_row
    for i in range(last_row,1,-1):
        cell = f"B{i}"
        value = ws[cell].value
        if value != None:
            if is_date(ws[cell].value):
                return False
    return True

def get_last_month_money_message(wb, month, year):
    if(month == "1" or month == "01"):
        try:
            last_year = int(year) - 1 
            last_year_wb = load_workbook(os.path.join(OUTPUT_FOLDER_PATH, str(last_year),"money.xlsx"))
            last_month_ws = last_year_wb["Dec"]

        except FileNotFoundError:
            return None
    else:
        try:
            last_month_ws = wb[month_to_string(int(month)-1)]
        except KeyError:
            return None

    ws_last_row = last_month_ws.max_row
    cell = f"B{ws_last_row}"
    ws_last_row_colB_value = last_month_ws[cell].value
    if(ws_last_row_colB_value != "current have"):
        return None
    else:
        result = []
        for col in [3,5,7,9,11,13,19]:
            cell = f"{get_column_letter(col)}{ws_last_row}"
            value = last_month_ws[cell].value
            result.append(value)
        return result 


def cell_type_message(ws,item): # item = [ [row(int), col(int), value(String)],[row1, col1, value1],...   ]
    for col, row, message in item:
        cell = f"{get_column_letter(col)}{row}"
        ws[cell].value = message
        ws[cell].font = helper.Font.FONT

def range_cell_set_color(ws, item): # item = [ [col1,row1,col2,row2,color1],[col1,row1,col2,row2,color]]
    for col1, row1, col2, row2, color in item:
        for row in ws.iter_rows(min_col = col1 , min_row = row1,
                                max_col = col2 , max_row = row2):
            for cell in row:
                cell.fill = color

def cell_set_border(ws, item): # item = [ [col1, row1, border1], [col2, row2, border2], ...]
    for col, row, border in item:
        cell = f"{get_column_letter(col)}{row}"
        ws[cell].border = border 

def range_merge_cell(ws, item): # item = [ [[col1, row1 (start cell)],[col2, row2 (end cell)]], ... ]
    for start , end in item:
        start_cell = f"{get_column_letter(start[0])}{start[1]}"
        end_cell = f"{get_column_letter(end[0])}{end[1]}"
        ws.merge_cells(f"{start_cell}:{end_cell}")
        ws[start_cell].alignment = Alignment(horizontal='center', vertical='top')


def sheet_init(wb, ws, month, year): # funciton for the first day of every month

    ws.column_dimensions['B'].width = 12

    for i in range(3,21):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = 10.5

    last_month_have = get_last_month_money_message(wb, month, year)

    cell_message = [
        [2, 2, "Date"], [3 , 2, "Cash"], [5, 2, "HSBC"], [7, 2, "Octopus"], [9, 2, "AliPay"], [11, 2, "Tap and Go"], [13, 2, "Payme"], [ 15, 2, "Countable Total"], [17, 2, "Real Total"], [20, 2, "Correct"], 
        [2, 3, "DD/MM/YY"], [3, 3, 'I'], [4, 3, 'O'], [5, 3, 'I'], [6, 3, 'O'], [7, 3, 'I'], [8, 3, 'O'], [9, 3, 'I'], [10, 3, 'O'], [11, 3, 'I'], [12, 3, 'O'], [13, 3, 'I'], [14, 3, 'O'], [15, 3, 'I'], [16, 3, 'O'], [17, 3, 'I'], [18, 3, 'O'], [19, 3, "Have"]
    ]

    if(last_month_have != None):
        last_message =[2, 4, "Last"], [3, 4, last_month_have[0]], [5, 4,last_month_have[1]], [7, 4,last_month_have[2] ], [9, 4,last_month_have[3] ], [11, 4,last_month_have[4] ], [13, 4,last_month_have[5] ], [ 19, 4,last_month_have[6] ]
    else:
        last_message= [2, 4, "Last"], [3, 4, "None"], [5, 4, "None"], [7, 4, "None"], [9, 4, "None"], [11, 4, "None"], [13, 4, "None"], [ 19, 4, "None"]
    
    for message in last_message:
        cell_message.append(message)

    cell_color = [
        [3, 2, 4, 4, helper.Color.PALE_MINT],
        [5, 2, 6, 4, helper.Color.PALE_PINK],
        [7, 2, 8, 4, helper.Color.LIGHT_BLUE],
        [9, 2, 10, 4, helper.Color.SKY_BLUE],
        [11, 2, 12, 4, helper.Color.PEACH],
        [13, 2, 14, 4, helper.Color.LACENDER],
        [15, 2, 16, 4, helper.Color.BEIGE],
        [17, 2, 19, 4, helper.Color.LIGHT_CORAL],
        [20, 2, 20, 4, helper.Color.YELLOW_GREEN],
    ]

    cell_border = [ 
        *([2+i, 2 , helper.Border.F_M ] for i in range(0,19)),
        *([2+i, 2+j , helper.Border.F_M] for i in [0,18] for j in [1,2]),
        *([2+i, 2+j , helper.Border.TBL_M_R_D] for i in range(1,16) if i%2==1 for j in [1,2]),
        *([2+i, 2+j , helper.Border.TBR_M_L_D] for i in range(1,16) if i%2==0 for j in [1,2]),
        *([18, 2+j, helper.Border.TB_M_LR_D] for j in [1,2]),
        *([19, 2+j, helper.Border.TBR_M_L_D] for j in [1,2]),
    ] 

    merge_cell = [
        [[3, 2], [4, 2]],
        [[5, 2], [6, 2]],
        [[7, 2], [8, 2]],
        [[9, 2], [10, 2]],
        [[11, 2], [12, 2]],
        [[13, 2], [14, 2]],
        [[15, 2], [16, 2]],
        [[17, 2], [19, 2]]
    ]

    cell_type_message(ws,cell_message)
    range_cell_set_color(ws, cell_color)
    cell_set_border(ws, cell_border)
    range_merge_cell(ws, merge_cell)

def is_date_duplicate(ws,date):
    last_row = ws.max_row
    for row in range(last_row, 1,-1):
        cell = f"B{row}"
        value = ws[cell].value
        if value != None:
            if is_date(value):
                if (value == date):
                    return True
                else:
                    return False            

def find_start_row(ws):
    last_row = ws.max_row

    for i in range(last_row, 1, -1):
        if any(cell.value is not None for cell in ws[i]):
            return i+1
    return 5

def input_date(ws, message):
    start_row = find_start_row(ws)
    cell_value = [[2,start_row, message]]
    cell_type_message(ws, cell_value)
    return start_row

ACCOUNT_LIST = [ ["C", 1], ["H", 3], ["O", 5], ["AP", 7], ["TnG", 9], ["P", 11]]

def self_convert(ws, amount, account, to, record_list):
    account_add_amount = next(num for key,num in ACCOUNT_LIST if key == account)
    to_add_amount = next(num for key,num in ACCOUNT_LIST if key == to)
    col, row  = record_list[0], record_list[1]
    message = [
        [col+account_add_amount, row, to],
        [col+account_add_amount+1, row, amount],
        [col+to_add_amount, row, amount],
        [col+to_add_amount+1, row, account]
    ]
    cell_type_message(ws, message)

def income(ws, amount, account, to, countable, record_list):
    to_add_amount = next(num for key, num in ACCOUNT_LIST if key == to)  
    col , row = record_list[0],record_list[1]
    message = [
        [col+to_add_amount, row, amount],
        [col+to_add_amount+1, row, account],
        [col+13, row , amount],
        [col+14, row, account],
        [col+15, row, amount],
        [col+16, row, account]
    ]
    if(not countable):
        message.pop(2)
        message.pop(2) 
   
    cell_type_message(ws, message)
    

def pay(ws, amount, account, to, countable, record_list):
    account_add_amount = next(num for key,num in ACCOUNT_LIST if key == account)
    col, row = record_list[0], record_list[1]
    message = [
        [col+account_add_amount, row, to],
        [col+account_add_amount+1, row, amount],
        [col+13, row, to],
        [col+14, row, amount],
        [col+15, row, to],
        [col+16, row, amount]
    ]     
    if(not countable):
        message.pop(2)
        message.pop(2)

    cell_type_message(ws, message)

def find_last_value_row(ws, value, row):
    for i in range(row,3,-1):
        cell = f"B{i}" 
        if (ws[cell].value == value):
            return i
    return None

def get_last_day(day):
    day = datetime.strptime(day, "%d/%m/%Y")
    return (day - timedelta(days=1)).strftime("%d/%m/%Y")

def is_num(num):
    try:
        float(num)  
        return True
    except ValueError:
        return False

def cal_sum_in_outcome(ws, col, start_row, end_row):
    return_value = Decimal(0);
    for row in range(start_row, end_row):

        cell = f"{get_column_letter(col)}{row}"
        try:
            cell_value = ws[cell].value
            if(is_num(cell_value)):
                cal_value =  Decimal(cell_value)
            else:
                cal_value = Decimal('0')

        except TypeError:
            cal_value =  Decimal('0')
        return_value += cal_value
    return return_value

def day_summary(ws, date, end_row):
    date_start_row = find_last_value_row(ws, date, end_row)

    day , month, year = date.split('/')

    if (date_start_row == None):
        message_out(f"Money Excel: Error, cannot find last cell of value: {date}")

    if(day == "01"):
        last_row = find_last_value_row(ws, "Last", end_row)
    else:
        last_row = find_last_value_row(ws, "sum", end_row)

    total_money = Decimal('0')
    for pair in ACCOUNT_LIST:
        income_money = Decimal('0')
        outcome_money = Decimal('0')

        col = 2 + pair[1]

        last_cell = f"{get_column_letter(col)}{last_row}"
        last_money = Decimal(ws[last_cell].value) 

        income_money = cal_sum_in_outcome(ws, col, date_start_row, end_row) 
        outcome_money = cal_sum_in_outcome(ws, col+1, date_start_row, end_row)

        total_value = last_money +  income_money - outcome_money 
        total_money += total_value

        message = [ [2,end_row, "sum"]]
        cell_type_message(ws, message)

        message = [[col, end_row, total_value]]

        cell_type_message(ws,message)

    message = [[19, end_row, total_money]]
    cell_type_message(ws, message)

    for col in [15,17]:
        total_input = cal_sum_in_outcome(ws, col , date_start_row, end_row)
        total_output = cal_sum_in_outcome(ws, col+1 , date_start_row, end_row)

        message = [[col, end_row, total_input],[col+1, end_row, total_output]]

        cell_type_message(ws, message)

    last_have_value = Decimal(str(ws[f"{get_column_letter(19)}{last_row}"].value)).quantize(Decimal("0.00"))
    current_input_value = Decimal(str(ws[f"{get_column_letter(17)}{end_row}"].value)).quantize(Decimal("0.00"))
    current_output_value = Decimal(str(ws[f"{get_column_letter(18)}{end_row}"].value)).quantize(Decimal("0.00"))
    current_have_value = Decimal(str(ws[f"{get_column_letter(19)}{end_row}"].value)).quantize(Decimal("0.00"))
    
    correct_value = str(last_have_value == current_have_value + current_output_value - current_input_value)
    message = [[20, end_row, correct_value]]
    cell_type_message(ws, message)


def day_total(ws, date, row):
    day, month, year = date.split('/')
    current_real_sum_in_cell = f"Q{row-1}"
    current_real_sum_out_cell = f"R{row-1}"
    current_real_sum_in_value = Decimal(str(ws[current_real_sum_in_cell].value))
    current_real_sum_out_value = Decimal(str(ws[current_real_sum_out_cell].value))

    current_countable_sum_in_cell = f"O{row-1}"
    current_countable_sum_out_cell = f"P{row-1}"
    current_countable_sum_in_value = Decimal(str(ws[current_countable_sum_in_cell].value))
    current_countable_sum_out_value = Decimal(str(ws[current_countable_sum_out_cell].value))

    if(day == "01"):
        message = [[2, row, "total"], [15, row, current_countable_sum_in_value], [16, row, current_countable_sum_out_value],[17, row, current_real_sum_in_value], [18, row, current_real_sum_out_value]] 
    else:
        last_row = find_last_value_row(ws, "total", row)
        last_real_total_in_cell = f"Q{last_row}" 
        last_real_total_out_cell = f"R{last_row}"
        last_real_total_in_value = Decimal(str(ws[last_real_total_in_cell].value))
        last_real_total_out_value = Decimal(str(ws[last_real_total_out_cell].value))

        last_countable_sum_in_cell = f"O{last_row}"
        last_countable_sum_out_cell = f"P{last_row}"
        last_countable_sum_in_value = Decimal(str(ws[last_countable_sum_in_cell].value))
        last_countable_sum_out_value = Decimal(str(ws[last_countable_sum_out_cell].value))

        total_in_value = last_real_total_in_value + current_real_sum_in_value
        total_out_value = last_real_total_out_value + current_real_sum_out_value

        countable_total_in_value = last_countable_sum_in_value + current_countable_sum_in_value
        countalbe_total_out_value = last_countable_sum_out_value + current_countable_sum_out_value


        message = [[2, row, "total"], [15, row, countable_total_in_value], [16, row, countalbe_total_out_value], [17, row, total_in_value], [18, row, total_out_value]]

    cell_type_message(ws, message)

def mark_current_amount(ws, date, row):
    if(is_folder_exist(CURRENT_HAVE_FILE_PATH)):
        full_file = read_file_data(CURRENT_HAVE_FILE_PATH)

        message = [[2, row, "current have"]]
        is_current_day = False
        total_amount = Decimal("0.00")
        for line in full_file:
            if(line == date):
                is_current_day = True
                continue
            elif(is_current_day):
                account , amount = line.split(' ')
                amount = Decimal(amount)
                col = next(num for key, num in ACCOUNT_LIST if account == key)

                sum_cell = f"{get_column_letter(2+col)}{row-2}"
                sum_value = Decimal(str(ws[sum_cell].value)).quantize(Decimal("0.00"))

                total_amount += amount

                correct_value = "True" if (sum_value == amount) else "False"

                message.append([2+col, row, amount])
                message.append([col+3, row, correct_value]) 
            
            else:
                return False

        message.append([19, row, total_amount]) 
        cell_type_message(ws, message)
        return True

    else:
        return False

def day_set_style(ws, date, row, current_marked):
    date_row = find_last_value_row(ws, date, row)
    cell_color= []
    if(current_marked):    
        # row = sum row, row + 1 = total row, row + 2 = current have row
        cell_border = [
            *([i, date_row, helper.Border.TLR_M_B_D] for i in [2,20]),
            *([i, date_row, helper.Border.TL_M_RB_D] for i in range(3,18) if i%2==1),
            *([i, date_row, helper.Border.TR_M_LB_D] for i in range(4,19) if i%2==0),
            [18,date_row , helper.Border.T_M_BLR_D],

            *([i, j, helper.Border.TB_D_LR_M] for i in range [2,20] for j in range(date_row+1,row-1)), # only for the first row middle part
            *([i ,j, helper.Border.TRB_D_L_M] for i in range(3,18) if i%2==1 for j in range(date_row+1, row-1)),
            *([i, j, helper.Border.TLB_D_R_M] for i in range(4,20) if (i%2==0 and i==19) for j in range(date_row+1, row-1)),
            *([18,j, helper.Border.F_D] for j in range(date_row+1, row-1)),

            *([i,row-1, helper.Border.T_D_LR_M_B_T] for i in range [2,20]),
            *([i, row-1, helper.Border.TR_D_L_M_B_M])


            # flag change the border method ( T, R, B , L) by border style
        ]
# cell_border = [ 
#         *([2+i, 2 , helper.Border.F_M ] for i in range(0,19)),
#         *([2+i, 2+j , helper.Border.F_M] for i in [0,18] for j in [1,2]),
#         *([2+i, 2+j , helper.Border.TBL_M_R_D] for i in range(1,16) if i%2==1 for j in [1,2]),
#         *([2+i, 2+j , helper.Border.TBR_M_L_D] for i in range(1,16) if i%2==0 for j in [1,2]),
#         *([18, 2+j, helper.Border.TB_M_LR_D] for j in [1,2]),
#         *([19, 2+j, helper.Border.TBR_M_L_D] for j in [1,2]),
#     ]
    range_cell_set_color(ws, cell_color) 
    cell_set_border(ws,cell_border)

def money_excel_process():
    message_out("Money excel processing")

    # --- copy to here --- #    

def day_finish(ws, date, row):
    day_summary(ws, date, row)
    day_total(ws, date, row+1)
    current_marked = mark_current_amount(ws, date, row+2)
    day_set_style(ws, date, row, current_marked)

def test():
    # --- need to copy to the process function --- #

    file_first_line = read_first_line_of_file(MONEY_MESSAGE_FILE_PATH)
    if(is_date(file_first_line)):
        f_day, f_month, f_year = file_first_line.split('/')
        MONEY_FOLDER_PATH = os.path.join(OUTPUT_FOLDER_PATH, f_year)
        MONEY_FILE_PATH = os.path.join(MONEY_FOLDER_PATH, "money.xlsx")
    else:
        if(message_out):
            message_out("Money Excel : Error, the money_messages.txt first is not a dat!") 
            sys.exit(0)

    if(not is_folder_exist(MONEY_FOLDER_PATH)):
        os.makedirs(MONEY_FOLDER_PATH)

    wb = check_and_open_excel(MONEY_FILE_PATH)
    ws = wb.active

    current_date = None
    money_message_list = read_file_data(MONEY_MESSAGE_FILE_PATH)
    START_COL = 2
    have_previous_date = False
    for message in money_message_list:
        if is_date(message):
            day, month, year = message.split('/') 
            month_str = month_to_string(month)
            ws = open_month_ws(wb, month_str)

            current_date = message

            if(have_previous_date):
                last_day = get_last_day(message)
                day_finish(ws, last_day, start_row + record)
                

            if(is_first_time_of_ws(ws)): 
                sheet_init(wb, ws, month, year)
            else:
                if(is_date_duplicate(ws,message)):
                    message_out(f"Money_excel : date {message} is marked!")
                    sys.exit(0) # for just in testing, may need to change after apply to gui

            start_row = input_date(ws,message)
            record = 0
        else: 
            have_previous_date = True
            amount, account, to, countable, *rest = shlex.split(message)  # *rest is just let the code can run successfully 
            account = account.strip('"')
            to = to.strip('"')
            countable = countable == "True"
            if any(to == pair[0] for pair in ACCOUNT_LIST) and any(account == pair[0] for pair in ACCOUNT_LIST ):
                self_convert(ws, amount, account, to, [START_COL, start_row + record])
            elif any(to == pair[0] for pair in ACCOUNT_LIST):
                income(ws,amount, account, to , countable, [START_COL, start_row + record])
            else:
                pay(ws, amount, account, to, countable, [START_COL, start_row + record])
            record+=1

    if(current_date is not None):
        day_finish(ws, current_date, start_row + record)

    wb.save(MONEY_FILE_PATH)

    # --- stop copy --- #

if __name__ == "__main__":
    test()