from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook

year = 2019 
registernum = '*64697\-16*'
rows = 2
start = 0
count = 0
sim = "'"
RegisterNumber = []
RegisterModification = []
SerialNumber = []
ModelName = []
totalChar = []
"""
# url = 'https://fgis.gost.ru/fundmetrology/cm/results?filter_mi_mitnumber=64697-16&activeYear=2018'
# url = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:2018&fq=mi.mitnumber:*64697\-16*&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows=100&start=0'
# response = requests.get(url)
# print(response)
# bs4 = BeautifulSoup(response.content,"lxml")
url2 = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:'+str(year)+'&fq=mi.mitnumber:'+registernum+'&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows='+str(rows)+'&start='+str(start)
response = requests.get(url2) #происходит попытка извлечь данные из определенного ресурса
#print(response)
bs4 = BeautifulSoup(response.content,"lxml") #Создается объект BeautifulSoup, HTML-данные передаются конструктору. 
#Второй параметр определяет синтаксический анализатор.
#print(bs4)
PageData = bs4.get_text() #возвращает весь текст HTML-документа или HTML-тега в виде единственной строки
"""


def receive_page(year, registernum, rows, start):
    url2 = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:' + str(
        year) + '&fq=mi.mitnumber:' + registernum + '&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows=' + str(
        rows) + '&start=' + str(start)
    response = requests.get(url2)  # происходит попытка извлечь данные из определенного ресурса
    # print(response)
    bs4 = BeautifulSoup(response.content, "lxml")  # Создается объект BeautifulSoup, HTML-данные передаются конструктору.
    # Второй параметр определяет синтаксический анализатор.
    # print(bs4)
    PageData = bs4.get_text()  # возвращает весь текст HTML-документа или HTML-тега в виде единственной строки
    return PageData


PageData = receive_page(year, registernum, rows, start)


def find_max_range_index(PageData):
    IndexMaxRange = PageData.find('"numFound":')
    IndexMaxRange2 = PageData.find(',', IndexMaxRange + 11, IndexMaxRange + 17)
    CountMaxRange = int(PageData[IndexMaxRange + 11:IndexMaxRange2])  # срез
    return CountMaxRange
    # print(CountMaxRange)  # 235 - количество всех строк


"""
#Поиск максимального значения индекса в год
IndexMaxRange = PageData.find('"numFound":')
IndexMaxRange2 = PageData.find(',',IndexMaxRange+11,IndexMaxRange+17)
CountMaxRange = int(PageData[IndexMaxRange+11:IndexMaxRange2]) #срез 
print(CountMaxRange) #235 - количество всех строк 
"""
Count_Max_Range = find_max_range_index(PageData)


def find_info(index, ResultString, index_plus, index_plus_end):
    if index == -1:
        result = ' '
        # RegisterNumber.append(' ')
        # return RegisterNumber
        return result
    else:
        Index2 = ResultString.find(',', index + index_plus, index + index_plus_end)
        # print(IndexRegisterNumber)
        result = ResultString[index + index_plus:Index2 - 1]
        # RegisterNumber.append(ResultString[index + 16:Index2 - 1])
        # print(ResultString[IndexRegisterNumber+16:IndexRegisterNumber2-1])
        # return RegisterNumber
        return result


for i in range(Count_Max_Range // 100 + 1):
    rows = 100
    PageData = receive_page(year, registernum, rows, start)
    """
    url2 = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:'+str(year)+'&fq=mi.mitnumber:'+registernum+'&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows='+str(rows)+'&start='+str(start)
    response = requests.get(url2)
    bs4 = BeautifulSoup(response.content,"lxml")
    PageData = bs4.get_text()
    """

    IndexDict = PageData.find('"docs":')
    Result = PageData[IndexDict+7:-3]
    ResultList = Result.split("}") #получили список


    # NextRegisterNumber, NextModelName, NextSerialNumber, NextRegisterModification = 0, 0, 0, 0
    for j in range(rows):
        count += 1
        ResultString = ResultList[j]
        # print(ResultString)
        # print(count)


        # Запоминаем регистрационный номер в реестре
        IndexRegisterNumber = ResultString.find('"mi.mitnumber":"')
        RegisterNumber.append(find_info(IndexRegisterNumber, ResultString, 16, 28))
        # RegisterNumber = find_info(IndexRegisterNumber, ResultString)
        """
        # print(IndexRegisterNumber)
        if IndexRegisterNumber == -1:
            RegisterNumber.append(' ') 
        else:
            IndexRegisterNumber2 = ResultString.find(',',IndexRegisterNumber+16,IndexRegisterNumber+28)
            #print(IndexRegisterNumber)
            RegisterNumber.append(ResultString[IndexRegisterNumber+16:IndexRegisterNumber2-1]) 
            #print(ResultString[IndexRegisterNumber+16:IndexRegisterNumber2-1])
        """

        #Запоминаем модификацию устройства
        IndexRegisterModification = ResultString.find('"mi.modification":"')
        RegisterModification.append(find_info(IndexRegisterModification, ResultString, 19, 28))
        # RegisterModification = find_info(IndexRegisterModification, ResultString)
        """
        if IndexRegisterModification == -1:
            RegisterModification.append(' ')
        else:
            IndexRegisterModification2 = ResultString.find(',',IndexRegisterModification+19,IndexRegisterModification+28)
            #print(IndexRegisterModification)
            RegisterModification.append(ResultString[IndexRegisterModification+19:IndexRegisterModification2-1])
            #print(PageData[IndexRegisterModification+19:IndexRegisterModification2-1])
        """

        #Запоминаем серийный номер устройства
        IndexSerialNumber = ResultString.find('"mi.number":"')
        SerialNumber.append(find_info(IndexSerialNumber, ResultString, 13, 25))
        # SerialNumber = find_info(IndexSerialNumber, ResultString)
        """
        if IndexSerialNumber == -1:
            SerialNumber.append(' ')
        else:
            IndexSerialNumber2 = ResultString.find(',',IndexSerialNumber+13,IndexSerialNumber+25)
            #print(IndexSerialNumber)
            SerialNumber.append(ResultString[IndexSerialNumber+13:IndexSerialNumber2-1])
            #print(SerialNumber[i])
        """

        #Запоминаем имя модели
        IndexModelName = ResultString.find('"mi.mitype":"')
        ModelName.append(find_info(IndexModelName, ResultString, 13, 20))
        # ModelName = find_info(IndexModelName, ResultString)
        """
        if IndexModelName == -1:
            ModelName.append(' ')
        else:
            IndexModelName2 = ResultString.find(',',IndexModelName+13,IndexModelName+20)
            #print(IndexModelName)
            ModelName.append(ResultString[IndexModelName+13:IndexModelName2-1])
            #print(ModelName[i])
        """
        if count == Count_Max_Range:
            break

    start += 100


wb = Workbook()   # создаётся файл Excel
dest_filename = 'try.xlsx' # имя созданного файла Excel
sheet = wb.active
sheet.title = "static" # указывавем имя листу, с которым работаем
for row in range(1, 2):
    for i in range(Count_Max_Range):
        sheet.append([RegisterNumber[i], RegisterModification[i], SerialNumber[i], ModelName[i]])
wb.save(dest_filename) # сохраняем файл, с которым работали
print('Файл создан')
