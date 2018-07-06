import datetime
import requests
import pandas
import json

# Get RT 5-minute LMPs for all plants from selected date
acceptableDate = False
while acceptableDate is False:
    year = input('Select a year: [yyyy] ')
    month = input('Select a month: [mm] ')
    day = input('Select a day: [dd] ')

    if datetime.date(int(year), int(month), int(day)) < (datetime.date.today() + datetime.timedelta(days=2)):
        acceptableDate = True
    else:
        print('')
        print('Unacceptable date')

date = datetime.date(int(year), int(month), int(day))
date0 = datetime.date(int(year), int(month), int(day)).strftime('%Y-%m-%dT00:00:00')

API_KEY = 'f6155aeeff864197a6a0aa1aec1af5fe'

headers = {'Content-Type': 'application/x-www-form-urlencoded',
           'Ocp-Apim-Subscription-Key': API_KEY}

pnodes = ['51288','8445784','31065427','32418343','32418345','32418347','32418349','32418435','32418437','32418447','32418449','32418451','32418455','32418457','32418459','32418503','32418505','32418559','32418561','32418563','32418611','32419345','32419347','32419349','34497125','34497127','34497179','34497181','34508503','34509943','34509945','34509949','35079887','35079889','38367953','40243775','40243777','40243779','40243781','40243783','40243785','40243787','40243789','40243797','40243799','40243801','40243803','40243805','40243807','40243809','40243811','40243813','40243819','40243821','40243823','40243831','40243833','40243835','40243837','40243839','40243841','40243843','40243845','40243847','40243849','40243851','40243861','40243863','40243865','40243867','40243869','40243871','40243873','40243881','40243883','40243885','40243887','40523629','47330703','115944303','116013753','116472937','124076095','1069452904','1258625176','1269364670','1269364671','1269364672','1269364674','1338431126','1338431127','1338431128','1338431129','1338431130','1338431131','1338431132','1338431133','1338431134','1338431135','1338431136','1379659480','1379659500']

firstRun = True

for plant in pnodes:

    print("Pulling LMPs for " + str(plant))

    url = 'https://api.pjm.com/api/v1/da_hrl_lmps?download=false&rowCount=50000&startRow=1&datetime_beginning_ept=' + str(date0) + '&pnode_id=' + plant
    response = requests.request("GET", url, headers=headers)
    result = json.loads(response.content)['items']
    result = pandas.DataFrame(result)

    if firstRun:
        data = result
        firstRun = False
    else:
        data = pandas.concat([data, result])

# Change date format
print('')
print('Updating date formats')

def convert_to_date(s):
    s = s.split('T')
    s0 = s[0].split('-')
    s1 = s[1]
    sr = str(s0[1]) + str(s0[2]) + str(s0[0]) + str(s1)

    return sr

data['datetime_beginning_utc'].apply(convert_to_date)
data['datetime_beginning_ept'].apply(convert_to_date)

# Re-arrange to proper format
data = data.assign(pnode_name=0)
data = data[['datetime_beginning_utc', 'datetime_beginning_ept', 'pnode_id', 'pnode_name', 'voltage', 'equipment', 'type', 'zone', 'system_energy_price_da', 'total_lmp_da', 'congestion_price_da', 'marginal_loss_price_da', 'row_is_current', 'version_nbr']]

# Write daily file
saveLoc = 'L:/Resource Scheduling/Conor/Data/ftr_lmps/'

data.to_csv(saveLoc + str(date.strftime('%Y%m%d')) + '_da_lmps.csv', index=False)

print('')
print('Successfully pulled LMPs for ' + str(date.strftime('%Y-%m-%d')))
print('File saved at ' + str(saveLoc))