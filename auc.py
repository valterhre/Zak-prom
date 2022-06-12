import math
import os
import requests
from openpyxl import Workbook

stores=['482010105', '48201029', '48201030', '48201031', '48201036', '48201070', '48210023', '48215610', '48215611', '48215633', '48225531', '48241094', '48246401', '48246403', '48246407', '48246414', '48246415', '48246418', '48246423', '48246424', '48246425', '48250029', '48257001', '48267602', '48277602', '48280099', '48280187', '48280214']
names={'482010105':'NOVUS SkyMall',
'48201029':'NOVUS Кольцевая',
'48201030':'NOVUS Киев Левобережная',
'48201031':'NOVUS Retroville',
'48201036':'NOVUS Бажана DRIVE',
'48201070':'NOVUS Осокор',
'48210023':'Пчелка Отрадный DRIVE',
'48215610':'METRO Григоренко DRIVE',
'48215611':'METRO Теремки DRIVE',
'48215633':'METRO Троещина DRIVE',
'48225531':'Космос Киев DRIVE',
'48241094':'Varus Вышгородская',
'48246401':'Auchan Петровка',
'48246403':'Auchan Кольцевая',
'48246407':'Auchan Беличи DRIVE',
'48246414':'Auchan Rive Gauche',
'48246415':'Auchan Лыбедская',
'48246418':'Auchan Черниговская DRIVE',
'48246423':'Auchan Глушкова',
'48246424':'Auchan Семьи Сосниных',
'48246425':'Auchan Луговая',
'48250029':'СитиМаркет Гостомель DRIVE',
'48257001':'Столичный DRIVE',
'48267602':'MEGAMARKET Kosmopolit DRIVE',
'48277602':'Ultramarket Kosmopolit DRIVE',
'48280099':'ЕкоМаркет Борисполь Киевский путь DRIVE',
'48280187':'ЕкоМаркет Жилянская',
'48280214':'ЕкоМаркет Закревского'}
url='https://stores-api.zakaz.ua/stores/{}/products/promotion/?page={}'
wb=Workbook()
ws=wb.active
li=[]
ws.append(['Title', 'Price','value of discount', "Due-datw",'Store-name' ])
headers = {'accept-language': 'ru,en-US;q=0.9,en;q=0.8,uk;q=0.7'}
try:
    for store in stores:
        url1 = url.format(store, 1)
        res = requests.get(url1).json()
        for k in range(1, math.ceil(res['count']/30)+1):
            url1=url.format(store, k)
            res = requests.get(url1, headers=headers).json()
            for m in range(len(res['results'])):
                title=res['results'][m]['title']
                price=res['results'][m]['price']/100
                due_date=res['results'][m]["discount"]["due_date"]
                value=res['results'][m]["discount"]["value"]
                ws.append([title,price,value, due_date,names[store]])
                wb.save('promotion.xlsx')
except:
    print('KeyError: "count"')
wb.save('promotion1.xlsx')
os.startfile('promotion1.xlsx')
