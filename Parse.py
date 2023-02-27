from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
#from openpyxl.utils import get_column_letter #???
import time 
import datetime

year = 2019 
registernum = '*64697\-16*'
rows = 2
start = 0
count = 0
sim="'"
RegisterNumber = []
RegisterModification = []
SerialNumber = []
ModelName = []
totalChar = []
#url = 'https://fgis.gost.ru/fundmetrology/cm/results?filter_mi_mitnumber=64697-16&activeYear=2018'
#url = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:2018&fq=mi.mitnumber:*64697\-16*&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows=100&start=0'
#response = requests.get(url)
#print(response)
#bs4 = BeautifulSoup(response.content,"lxml")
url2 = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:'+str(year)+'&fq=mi.mitnumber:'+registernum+'&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows='+str(rows)+'&start='+str(start)
response = requests.get(url2) #происходит попытка извлечь данные из определенного ресурса
#print(response)
bs4 = BeautifulSoup(response.content,"lxml") #Создается объект BeautifulSoup, HTML-данные передаются конструктору. 
#Второй параметр определяет синтаксический анализатор.
#print(bs4)
PageData = bs4.get_text() #возвращает весь текст HTML-документа или HTML-тега в виде единственной строки

#Поиск максимального значения индекса в год
IndexMaxRange = PageData.find('"numFound":')
IndexMaxRange2 = PageData.find(',',IndexMaxRange+11,IndexMaxRange+17)
CountMaxRange = int(PageData[IndexMaxRange+11:IndexMaxRange2]) #срез 
print(CountMaxRange) #235 - количество всех строк 

for i in range (CountMaxRange):
    count += 1
    rows=i+1
    url2 = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:'+str(year)+'&fq=mi.mitnumber:'+registernum+'&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows='+str(rows)+'&start='+str(start)
    response = requests.get(url2)
    bs4 = BeautifulSoup(response.content,"lxml")
    PageData = bs4.get_text()
    #print(bs4)

   #Запоминаем регистрационный номер в реестре
    IndexRegisterNumber = PageData.rfind('"mi.mitnumber":"')
    #print(IndexRegisterNumber)
    IndexRegisterNumber2 = PageData.find(',',IndexRegisterNumber+16,IndexRegisterNumber+28)
    #print(IndexRegisterNumber)
    RegisterNumber.append(PageData[IndexRegisterNumber+16:IndexRegisterNumber2-1]) 
    #print(RegisterNumber[i])

   #Запоминаем модификацию устройства
    IndexRegisterModification = PageData.rfind('"mi.modification":"')
    IndexRegisterModification2 = PageData.find(',',IndexRegisterModification+19,IndexRegisterModification+28)
   #print(IndexRegisterModification)
    RegisterModification.append(PageData[IndexRegisterModification+19:IndexRegisterModification2-1])
    #print(RegisterModification[i])

   #Запоминаем серийный номер устройства
    IndexSerialNumber = PageData.rfind('"mi.number":"')
    IndexSerialNumber2 = PageData.find(',',IndexSerialNumber+13,IndexSerialNumber+25)
   #print(IndexSerialNumber)
    SerialNumber.append(PageData[IndexSerialNumber+13:IndexSerialNumber2-1])
    #print(SerialNumber[i])

   #Запоминаем имя модели
    IndexModelName = PageData.rfind('"mi.mitype":"')
    IndexModelName2 = PageData.find(',',IndexModelName+13,IndexModelName+20)
   #print(IndexModelName)
    ModelName.append(PageData[IndexModelName+13:IndexModelName2-1])
    #print(ModelName[i])

    if count % 100 == 0:
        print(count) 
        print(datetime.datetime.now().time())
        time.sleep(10) 


"""file = open("parse.txt", "w")
for i in range (1):
    totalChar.append(RegisterNumber[i])
    totalChar.append(' ')
    totalChar.append(RegisterModification[i])
    totalChar.append(' ')
    totalChar.append(SerialNumber[i])
    totalChar.append(' ')
    totalChar.append(ModelName[i])
    totalChar.append('\n') 
    for i in range(len(totalChar)):
        file.write(totalChar[i]) 
"""

wb = Workbook()   # создаётся файл Excel
dest_filename = 'try.xlsx' # имя созданного файла Excel
sheet = wb.active
sheet.title = "static" # указывавем имя листу, с которым работаем
for row in range(1, 2): 
    for i in range(CountMaxRange):
        sheet.append([RegisterNumber[i], RegisterModification[i], SerialNumber[i], ModelName[i]])
wb.save(dest_filename) # сохраняем файл, с которым работали
print('Конец, создан файл tey.xlsx с листом range names, где заполнены строки 1-39 порядковым номером столбца')

