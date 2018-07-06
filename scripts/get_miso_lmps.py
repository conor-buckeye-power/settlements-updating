from io import StringIO
import xlwings as xl
import datetime
import requests
import pandas

def update_miso_lmps():
    # GET USER INPUT AND ENSURE VALUES ARE INTEGERS
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
        # SET MONTH AND YEAR WITH DATE SET TO THE FIRST
        date = datetime.datetime(year, month, 1)

        # SET STORY COUNTY WORKBOOK PATH
        story_county_path = 'H:/PS/Power Generation/Settlements/Story County/LMP Data/'

        # ========== DAY-AHEAD ==========
        # OPEN STORY COUNTY WORKBOOK
        da_wb = xl.Book(story_county_path + str(date.strftime('%Y')) + '/' + str(date.strftime('%b.%y')) + '/FTR LMPs.xls')

        # FIND THE NEXT DATE NEEDED
        for i in range(1, 32):
            da_sht = da_wb.sheets[str(i)]
            if da_sht.range('A4').value == None:
                da_iter = i
                break

        # ADD DAY-AHEAD LMPs
        while True:
            da_date = datetime.datetime(year, month, da_iter)
            da_sht = da_wb.sheets[str(da_iter)]
            resp = requests.get('https://docs.misoenergy.org/marketreports/' + str(da_date.strftime('%Y%m%d')) + '_da_expost_lmp.csv')
            r = str(resp)

            if r == '<Response [404]>':
                break
            else:
                data = pandas.read_csv(StringIO(resp.text), skiprows=4)
                data = data[data.Node == 'ALTW.STORYBUCK']
                rows = list(data.loc[:, 'Node':'HE 24'].index)

                a = data.loc[rows[0], 'Node':'HE 24']
                b = data.loc[rows[1], 'Node':'HE 24']
                c = data.loc[rows[2], 'Node':'HE 24']

                da_sht.range('A1:A27').options(index=False, header=False).value = a
                da_sht.range('A31:A57').options(index=False, header=False).value = b
                da_sht.range('A61:A87').options(index=False, header=False).value = c

            da_iter+=1

        # SAVE AND CLOSE WORKBOOKS
        da_wb.save()
        da_wb.close()

        # ========== REAL-TIME ==========
        # OPEN STORY COUNTY WORKBOOK
        rt_wb = xl.Book(story_county_path + 'Real-Time/' + str(date.strftime('%Y')) + '/' + str(date.strftime('%b.%y')) + '/FTR LMPs.xls')

        # FIND THE NEXT DATE NEEDED
        for i in range(1, 32):
            rt_sht = rt_wb.sheets[str(i)]
            if rt_sht.range('A4').value == None:
                rt_iter = i
                break

        # ADD REAL-TIME LMPS
        while True:
            rt_date = datetime.datetime(year, month, rt_iter)
            rt_sht = rt_wb.sheets[str(rt_iter)]
            resp = requests.get('https://docs.misoenergy.org/marketreports/' + str(rt_date.strftime('%Y%m%d')) + '_rt_lmp_final.csv')
            r = str(resp)

            if r == '<Response [404]>':
                break
            else:
                data = pandas.read_csv(StringIO(resp.text), skiprows=4)
                data = data[data.Node == 'ALTW.STORYBUCK']
                rows = list(data.loc[:, 'Node':'HE 24'].index)

                a = data.loc[rows[0], 'Node':'HE 24']
                b = data.loc[rows[1], 'Node':'HE 24']
                c = data.loc[rows[2], 'Node':'HE 24']

                rt_sht.range('A1:A27').options(index=False, header=False).value = a
                rt_sht.range('A31:A57').options(index=False, header=False).value = b
                rt_sht.range('A61:A87').options(index=False, header=False).value = c

            rt_iter+=1

        # SAVE AND CLOSE WORKBOOKS
        rt_wb.save()
        rt_wb.close()

update_miso_lmps()



# ========== APPENDIX ==========
#rt_link = 'https://docs.misoenergy.org/marketreports/20180704_rt_lmp_final.csv'
#da_link = 'https://docs.misoenergy.org/marketreports/20180706_da_expost_lmp.csv'

#resp = requests.get(rt_link)
#data = pandas.read_csv(StringIO(resp.text), skiprows=4)

#data = data[data.Node == 'ALTW.STORYBUCK']
#rows = list(data.loc[:,'HE 1':'HE 24'].index)

#d = data.loc[rows[0], 'HE 1':'HE 24']
#d = list(d)

#da_sht.range('A4:A27').options(index=False).value = d
