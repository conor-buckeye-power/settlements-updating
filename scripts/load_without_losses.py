from calendar import monthrange
import xlwings as xl
import datetime
import pandas

def update_load_without_losses(year, month, exit_script):
    print('\n=============== LOAD WITHOUT LOSSES ===============')

    if exit_script is False:
        # Set path to load with and without losses folder
        path = 'L:/Resource Scheduling/Conor/Data/load_with_and_without_losses/'

        # Setup date format for the file format
        mranbg = str(year) + '-' + str(format(month, '02'))
        mrange = mranbg + '-01_' + mranbg + '-' + str(monthrange(year, month)[1]) + '.csv'

        # Read in the csv file
        try:
            data = pandas.read_csv(path + mrange, skiprows=4)

            # Filter only for load without losses, drop some unnecessary columns
            data = data[data.Status == 'Load without Losses']
            data = data.drop(['EPT HE 02*', 'Total MWh', 'Version'], axis=1)

            # Set the path to the Excel workbooks
            do_month = datetime.datetime(year, month, 1)
            write_path = 'H:/PS/Power Generation/Settlements/Load/Tracking Sheets/' + str(
                do_month.strftime('%b.%y')) + '/' + str(do_month.strftime('%b%y')) + ' Buck'

            # Set-up dictionary containing all zones to iterate over and their names from the report
            zones = {'AEP': 'AEPOPT',
                     'CINE': 'DEOEDC',
                     'DPL': 'DAYEDC',
                     'FE': 'OEEDC',
                     'IM': 'AEPIMP'}

            # Iterate over each zone, opening the respective workbook and add data
            for zone in zones:
                wb = xl.Book(write_path + zone + ' Loads.xlsx')
                sht = wb.sheets['Actuals']

                d = data[data.Buyer == zones[zone]]
                d = d.loc[:, 'EPT HE 01':'EPT HE 24']

                sht.range('B38:Y68').options(index=False, header=False).value = d

                wb.save()
                wb.close()

        except FileNotFoundError:
            print('ERROR! Load without Losses File Does Not Exist for ' + str(datetime.datetime(year, month, 1).strftime('%b-%Y')))


    '''
    book = 'H:/PS/Power Generation/Settlements/Load/Tracking Sheets/Jul.18/'

    wb = xl.Book(book + 'Jul18 BuckAEP Loads.xlsx')
    sht = wb.sheets['Actuals']
    sht.range('B38:Y68').options(index=False, header=False).value = data.loc[:, 'EPT HE 01':'EPT HE 24']
    '''

#update_load_without_losses(year, month, exit_script)