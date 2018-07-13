from calendar import monthrange
from openpyxl import load_workbook
import xlwings as xl
import datetime
import pandas

'''
Mone D50
GRNV B55
'''

def find_first_date(wb, cell):
    # Find the next date needed
    for i in range(1, 32):
        sht = wb.sheets[str(i)]
        if sht.range(cell).value == None:
            return(i)

def load_clearings(date):
    try:
        clearings = 'L:/Resource Scheduling/Conor/Clearings/' + str(date.strftime('%m.%d.%Y.xlsm'))
        clearings = load_workbook(clearings, data_only = True)
        return(clearings['Actual_Clearings'])
    except FileNotFoundError:
        return(404)

def update_da_clearings(year, month, exit_script):
    print('\n=============== DAY-AHEAD CLEARINGS ===============')

    # Cardinal
    card_path = 'H:/PS/Power Generation/Settlements/Cardinal/' +\
                str(datetime.datetime(year, month, 1).strftime('%Y/')) +\
                str(datetime.datetime(year, month, 1).strftime('%b%y')) +\
                ' Cardinal Settlement.xlsx'

    wb = xl.Book(card_path)
    day = find_first_date(wb, 'E30')

    for i in range(day, monthrange(year, month)[1]):
        sht = wb.sheets[str(i)]
        clr_ws = load_clearings((datetime.datetime(year, month, i) + datetime.timedelta(days=-1)))

        if clr_ws == 404:
            break
        else:
            for x in range(30, 54):
                valE = round(clr_ws['I' + str(x)].value, 0)
                valF = round(clr_ws['J' + str(x)].value, 0)
                sht.range('E' + str(x - 20)).value = -valE
                sht.range('F' + str(x - 20)).value = -valF

update_da_clearings(2018, 7, False)