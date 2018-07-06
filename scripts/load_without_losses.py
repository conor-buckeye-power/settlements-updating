import xlwings as xl
import pandas

#def update_load_without_losses():
path = 'L:/Resource Scheduling/Conor/Data/load_with_and_without_losses/'

data = pandas.read_csv(path + '2018-07-01_2018-07-31.csv', skiprows=4)
data = data[data.Status == 'Load without Losses']
data = data.drop(['EPT HE 02*', 'Total MWh', 'Version'], axis=1)



book = 'H:/PS/Power Generation/Settlements/Load/Tracking Sheets/Jul.18/'

wb = xl.Book(book + 'Jul18 BuckAEP Loads.xlsx')
sht = wb.sheets['Actuals']
sht.range('B38:Y68').options(index=False).value = data.loc[:, 'EPT HE 01':'EPT HE 24']