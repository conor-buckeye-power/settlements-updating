from xml.etree import ElementTree as ET
import pandas

path = 'H:/PS/Power Generation/Settlements/Story County/Settlement Data/2018/06-Jun 2018/RT_BUCK_07022018_06182018-S14.xml'
tree = ET.parse(path)

root = tree.getroot()
ASSET_OWNER = root.find('ASSET_OWNER')
ASSET = ASSET_OWNER.find('ASSET')
DET_TYP = ASSET.findall('DET_TYP')

dtn = []
val = []

for DT in DET_TYP:
    for element in DET_TYP:
        for name in element.findall('DET_TYP_NM'):
            for int_ in element.findall('INT'):
                dtn.append(name.text)
                val.append(int_.find('VAL').text)

data = pandas.DataFrame(data={'Type': dtn, 'Value': val})

d1 = data[data.Type == 'Excessive Energy Volume at Node: ALTW.STORYBUCK'].iloc[0:24]
d2 = data[data.Type == 'Real Time Metered Actual Volume at Node: ALTW.STORYBUCK'].iloc[0:24]
