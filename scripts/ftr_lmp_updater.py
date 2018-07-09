from calendar import monthrange
import xlwings as xl
import pywintypes
import datetime
import pandas
import os

def da_ftr_lmp_updater():
    # Get RT 5-minute LMPs for all plants from selected date
    exit_script = False
    try:
        year = int(input('Provide a year to update LMPs [yyyy]: '))
    except ValueError:
        print('\nERROR:   Values inputted may not be integers')
        print('WARNING: LMPs WILL NOT UPDATE\n')
        exit_script = True
    try:
        month = int(input('Provide a month to update LMPS  [mm]: '))
    except ValueError:
        print('\nERROR:   Values inputted may not be integers')
        print('WARNING: LMPs WILL NOT UPDATE\n')
        exit_script = True

    if exit_script is False:
        # Set-up dates
        date = datetime.date(int(year), int(month), 1)

        # Spreadsheet save location
        saveLoc = 'L:/Resource Scheduling/PJM Management/FTR LMPs/' + str(date.strftime('%b.%y')) + '/FTR LMPs.xls'

        # Open workbook for specified month
        wb = xl.Book(saveLoc)

        # Find the next date needed
        for i in range(1, 32):
            sht = wb.sheets[str(i)]
            if sht.range('A201').value == None:
                iter = i
                break

        # Find the end date
        beg_date = int(iter)
        end_date = int((datetime.datetime(year, month, monthrange(year, month)[1])).strftime('%d')) + 1

        for day in range(beg_date, end_date):
            sht = wb.sheets[str(day)]

            # Get data from Kevin's csv file
            file_in_loc = 'L:/Resource Scheduling/Kevin/LMP Data/DataMiner/'
            data_date = datetime.datetime(year, month, day)

            # Check to make sure input file exists
            path = file_in_loc + str(data_date.strftime('%Y%m%d_buckeye_nodes.csv'))
            if os.path.isfile(path) is True:
                # Get data from csv file
                data = pandas.read_csv(path)

                # Add data to worksheet
                sht.range('A201:N2648').options(index=False, header=False).value = data

                # Print success message
                print('Successfully added DA LMPs for ' + str(datetime.datetime(year, month, day).strftime('%Y-%m-%d')))
            else:
                break

        # Save and close workbook
        try:
            wb.save()
            wb.close()
        except pywintypes.com_error:
            print('WARNING! Exception occurred when saving Day-ahead FTR LMPs Workbook')

da_ftr_lmp_updater()

def rt_ftr_lmp_updater():
    pass

rt_ftr_lmp_updater()