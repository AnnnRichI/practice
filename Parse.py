from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import openpyxl
import time
import re

years = [2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023]
workbook_with_info_from_user = openpyxl.load_workbook("Вводные_данные.xlsx")
active_list_wb_from_user = workbook_with_info_from_user.active
year = active_list_wb_from_user['B7'].value
if year not in years:
    print("Неправильно указан год! Он должен быть в интервале с 2010 по 2023.\nИсправьте и запустите программу заново!)")
    exit()
year_end = active_list_wb_from_user['B9'].value
if year_end != None:
    if year_end not in years:
        print("Неправильно указан год! Он должен быть в интервале с 2010 по 2023.\nИсправьте и запустите программу заново!)")
        exit()
    if year > year_end:
        print("Второй год должен быть больше первого.\nИсправьте и запустите программу заново!)")
        exit()
else:
    year_end = year
registernum = str(active_list_wb_from_user['B12'].value)
check_registernum = re.match(r'\b\d{5}-\d{2}\b', registernum)
if check_registernum == None:
    print("Неправильно указан регистрационный номер.")
    exit()
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
range_for_print = []


def receive_page(year, registernum, rows, start):
    web_url = 'https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select?fq=verification_year:' + str(
        year) + '&fq=mi.mitnumber:' + registernum + '&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum,sticker_num&sort=verification_date+desc,org_title+asc&rows=' + str(
        rows) + '&start=' + str(start)
    response = requests.get(web_url)  # происходит попытка извлечь данные из определенного ресурса
    # print(response)
    bs4 = BeautifulSoup(response.content, "lxml")  # Создается объект BeautifulSoup, HTML-данные передаются конструктору.
    PageData = bs4.get_text()  # возвращает весь текст HTML-документа или HTML-тега в виде единственной строки
    # print(PageData)
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
    range_for_print.append(Count_Max_Range)
    all_rows += Count_Max_Range
    for i in range(Count_Max_Range // 100 + 1):
        rows = 100
        PageData = receive_page(year, registernum, rows, start)
        IndexDict = PageData.find('"docs":')
        Result = PageData[IndexDict+7:-3]
        ResultList = Result.split("}") #получили список

        for j in range(len(ResultList) - 1):
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

        start += 100
    year += 1

# print(range_for_print)
wb = Workbook()  # создаётся файл Excel
dest_filename = 'Result.xlsx'  # имя созданного файла Excel
sheet = wb.active
sheet.title = "static"  # указывавем имя листу, с которым работаем

j = 0
year_now = 0
year -= 1
for row in range(1, 2):
    sheet.append(['Регистрационный номер', 'Модификация', 'Заводской номер', 'Тип'])
    if len(range_for_print) == 1:
        sheet.append([year])
    for i in range(all_rows):
        if len(range_for_print) != 1:
            if year_now == 0 or i == year_now + range_for_print[j]:
                sheet.append([year + j])
                j += 1
                year_now = year_now + range_for_print[j]

        sheet.append([RegisterNumber[i], RegisterModification[i], SerialNumber[i], ModelName[i]])
wb.save(dest_filename)  # сохраняем файл, с которым работали
print('Файл создан')


