from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import openpyxl
import time

wb = openpyxl.load_workbook("input.xlsx")
ws = wb.active
year = ws['A1'].value
year_end = ws['B1'].value
registernum = str(ws['A2'].value)
registernum = f'{"*"}{registernum}{"*"}'

time.sleep(10)

rows = 2
start = 0
count = 0
sim = "'"
RegisterNumber = []
RegisterModification = []
SerialNumber = []
ModelName = []
totalChar = []
all_rows = 0


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


def find_max_range_index(PageData):
    IndexMaxRange = PageData.find('"numFound":')
    IndexMaxRange2 = PageData.find(',', IndexMaxRange + 11, IndexMaxRange + 17)
    CountMaxRange = int(PageData[IndexMaxRange + 11:IndexMaxRange2])  # срез
    return CountMaxRange
    # print(CountMaxRange)  # 235 - количество всех строк



def find_info(index, ResultString, index_plus, index_plus_end):
    if index == -1:
        result = ' '
        return result
    else:
        Index2 = ResultString.find(',', index + index_plus, index + index_plus_end)
        result = ResultString[index + index_plus:Index2 - 1]
        return result


while year != year_end + 1:
    start = 0
    count = 0
    PageData = receive_page(year, registernum, rows, start)
    Count_Max_Range = find_max_range_index(PageData)
    all_rows += Count_Max_Range
    for i in range(Count_Max_Range // 100 + 1):
        rows = 100
        PageData = receive_page(year, registernum, rows, start)
        IndexDict = PageData.find('"docs":')
        Result = PageData[IndexDict+7:-3]
        ResultList = Result.split("}") #получили список

        for j in range(rows):
            count += 1
            ResultString = ResultList[j]

            # Запоминаем регистрационный номер в реестре
            IndexRegisterNumber = ResultString.find('"mi.mitnumber":"')
            RegisterNumber.append(find_info(IndexRegisterNumber, ResultString, 16, 28))

            #Запоминаем модификацию устройства
            IndexRegisterModification = ResultString.find('"mi.modification":"')
            RegisterModification.append(find_info(IndexRegisterModification, ResultString, 19, 28))

            #Запоминаем серийный номер устройства
            IndexSerialNumber = ResultString.find('"mi.number":"')
            SerialNumber.append(find_info(IndexSerialNumber, ResultString, 13, 25))

            #Запоминаем имя модели
            IndexModelName = ResultString.find('"mi.mitype":"')
            ModelName.append(find_info(IndexModelName, ResultString, 13, 20))

            if count == Count_Max_Range:
                break

        start += 100
    year += 1


wb = Workbook()  # создаётся файл Excel
dest_filename = 'try.xlsx'  # имя созданного файла Excel
sheet = wb.active
sheet.title = "static"  # указывавем имя листу, с которым работаем
for row in range(1, 2):
    for i in range(all_rows):
        sheet.append([RegisterNumber[i], RegisterModification[i], SerialNumber[i], ModelName[i]])
wb.save(dest_filename)  # сохраняем файл, с которым работали
print('Файл создан')


