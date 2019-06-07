import os
from sys import exit
from openpyxl import Workbook, load_workbook
from easygui import diropenbox


#ссылка базовой папки
BASE_DIR = diropenbox()
#список файлов в указаной папке
folder_files_names = os.listdir(BASE_DIR)

#список подходящих файлов
xlsx_files_list = []
#список остальных файлов
not_xlsx_files_list = []
#список с данными с файлов
data_list = []
#файл с вкладками заявок
good_data_wb = Workbook()
#файл с вкладками другими
bad_data_wb = Workbook()
#ввести номер заявки
ask_number = int(input("Введи номер последней заявки > "))

def tuple_unzip(t):
    """
    Функция для распаковки екселевских тьюплов
    tuple format: ((value), )
    возвращает список с данными подобный тьюплу
    """
    l = []
    for i in t[0]:
        l.append(i.value)

    return l

def create_range(pos1, pos2, raw):
    """
    Превращает координаты ячеек екселя, в список имен ячеек.
    Нужна для указания места, куда ложить данные. 
    Принимает две буквенные координаты, и номер строки.
    pos1 - первая координата строки
    pos2 - вторая координата строки, может состоять из 2 букв
    raw - номер строки
    """

    if len(pos2) == 1:
        #если одна буква  
        ex_range = ['{0}{1}'.format(chr(i), raw) for i in range(
                        ord(pos1), ord(pos2)+1)]
        #если 2 буквы
    else:
        ex_range = ['{0}{1}'.format(chr(i), raw) for i in range(
                            ord(pos2[0]), 91)]
        double_range = ['{0}{1}{2}'.format('A', chr(i), raw) for i in range(
                            65, ord(pos2[1])+1)]
        ex_range.extend(double_range)

    return ex_range

        
for i in folder_files_names:
    if i.endswith('.xlsx'):
        xlsx_files_list.append(i)
    else:
        not_xlsx_files_list.append(i)
#проверяем не пуст ли список екселевских файлов
if xlsx_files_list:
    #проверяем екселевские файлы - заявки ли?
    for i in xlsx_files_list:
        #загружаем книги по очереди
        wb_temp = load_workbook(filename = os.path.join(BASE_DIR, i))
        #проверяем подходящие листи, проходим по листам в файле
        for j in wb_temp.sheetnames:
            if wb_temp[j]['A1'].value == "!!!":
                #если лист подходит, копируем информацию
                #строка
                raw_info = wb_temp[j]["A11":"AI11"]
                #фирма
                organization = wb_temp[j]["C12":"D12"]
                #мыло
                email = wb_temp[j]["F12":"G12"]
                #телефон
                phone = wb_temp[j]["B15"]
                #должность
                occupation = wb_temp[j]["G16":"I16"]
                #имя должостного лица
                name = wb_temp[j]["G17":"I17"]

                #формируем общий список -
                l = [raw_info, organization, email, phone, occupation, name]

                # добавляем в главный список
                data_list.append(l)                
#выходим с программы
else:
    print('В папке отсутствуют файлы ексель.')
    sys.exit()

#открываем шаблон для заявок
good_data_wb = load_workbook(filename = os.path.join(
                BASE_DIR, 'template.xlsx'))
ws = good_data_wb['л1']
for i in data_list:
    #номер с которого начнем
    ask_number += 1
    new_ws = good_data_wb.copy_worksheet(ws)
    new_ws.title = str(ask_number)



    """
                #создаем диапазон ячеек 
                raw_range = create_range('A','AI', 10)
                organization_range = create_range('C', 'D' , 11)
                email_range = create_range('F', 'G' , 11)
                phone_range = ["B14"]
                occupation_range = create_range('G', 'I' , 15)
                name_range = create_range('G', 'I' , 16)

                #распаковываем инфу
                raw_info_list =  tuple_unzip(raw_info)
                organization_list = tuple_unzip(organization)
                email_list = tuple_unzip(email)
                #phone_list = tuple_unzip(phone)
                phone_list = phone.value
                occupation_list = tuple_unzip(occupation)
                name_list = tuple_unzip(name)

                def put(range_l, info_l):
                    for i in range(len(range_l)):
                        good_data_wb['л1'][range_l[i]] = info_l[i]

                put(raw_range, raw_info_list)
                put(organization_range, organization_list)
                good_data_wb['л1'].merge_cells('C11:D11')
                put(email_range, email_list)
                good_data_wb['л1'].merge_cells('F11:H11')
                good_data_wb['л1'].print_area = 'A1:I16'
                good_data_wb['л1'].print_area = 'A1:AI16'



#загружаем книгу
#проверяем подходит ли нам эта книга? - та которую рассылал Валера? "!!!"
"""
good_data_wb.save('t1.xlsx') 