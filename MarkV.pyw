import os
import pyexcel
from pyexcel._compact import OrderedDict
import PySimpleGUI as sg
import subprocess
import shutil
import re

__progName__ = "MarkV" # Name of programm
__version__ = "2020_08_09" # Version of programm
sg.theme('System Default 1') # Theme: "Default 1", "DarkTanBlue", "System Default 1"...

minus = ("-", "−", "–")
hollow = ("", None)

# Загрузка файла параметров работы программы:
MarkVdatapath = os.path.join(os.path.split(__file__)[0], __progName__ + "data.xlsx")
MarkVdatapath = os.path.normpath(MarkVdatapath)
MarkVdata = pyexcel.get_book(file_name=MarkVdatapath)
sheet = MarkVdata["Трансфер"]
newsheetnames = sheet.row[0]
sheet.name_columns_by_row(0)
printprograms = MarkVdata["Программы"].row[0][1:]

importfilename = "ОБЪЕКТ.xls"

def data(page, key, value=None, book=MarkVdata):
    """ Возваращает словарь с ключами key и значениями value на листе page книги Excel
        key и value - заголовки столбцов листа page книги Excel
    """
    sheet = book[page]
    for column in sheet.columns():
        if column[0] == key:
            keys = column[1:]
        if column[0] == value:
            values = column[1:]
    if value == None:
        return keys
    else:
        dictionary = {}
        for i in range(len(keys)):
            dictionary[keys[i]] = values[i]
        return dictionary

def proginstalled(printprogram):
    if  os.path.exists(data("Программы", "KEY", printprogram)["programpath"]) and \
        os.path.exists(data("Программы", "KEY", printprogram)["packpathdst"]):
        return True
    else:
        return False
def packinstalled(printprogram):
    src = data("Программы", "KEY", printprogram)["packpathsrc"]
    dst = data("Программы", "KEY", printprogram)["packpathdst"]
    for srctuple in os.walk(src):
        for srcdir in srctuple[1]:
            path = os.path.join(srctuple[0], srcdir).replace(src, dst)
            # print(path)
            if not os.path.exists(path):
                return False
        for srcfile in srctuple[2]:
            path = os.path.join(srctuple[0], srcfile).replace(src, dst)
            # print(path)
            if not os.path.exists(path):
                return False
    return True
def installpack(printprogram):
    src = data("Программы", "KEY", printprogram)["packpathsrc"]
    dst = data("Программы", "KEY", printprogram)["packpathdst"]
    try:
        for srctuple in os.walk(src):
            for srcdir in srctuple[1]:
                path = os.path.join(srctuple[0], srcdir).replace(src, dst)
                if not os.path.exists(path):
                    print(path)
                    os.mkdir(path)
            for srcfile in srctuple[2]:
                srcfilepath = os.path.join(srctuple[0], srcfile)
                dstfilepath = srcfilepath.replace(src, dst)
                if not os.path.exists(dstfilepath):
                    print("Отсутствует файл:", dstfilepath)
                shutil.copyfile(srcfilepath, dstfilepath)
                print("Скопирован файл:", srcfilepath, "\nВ файл:", dstfilepath)
    except:
        return False
    else:
        return True

def preproc(book_dict):
    """ Предобработка таблиц данных: 
        удаление листов, начинающихся с "-", пустых строк, неозаглавленных столбцов, 
        заполнение ячеек, содержащих ссылку "^"
    """
    # Удаление листов с именами, начинающимися с минуса
    for sheetname in list(book_dict.keys()):
        if sheetname[0] in minus:
            del book_dict[sheetname]
            continue
        sheet_list = book_dict[sheetname]
        # Удаление строк содержащих только пустые или заполненные пробелами ячейки
        row_num = 0
        while row_num < len(sheet_list):
            row_empty = True
            for col_num in range(len(sheet_list[0])):
                if isinstance(sheet_list[row_num][col_num], str):
                    sheet_list[row_num][col_num] = sheet_list[row_num][col_num].strip()
                if sheet_list[row_num][col_num] not in hollow:
                    row_empty = False
                    break
            if row_empty:
                del sheet_list[row_num]
            else:
                row_num += 1
        # Удаление неозаглавленных столбцов данных:
        col_num = 0
        while col_num < len(sheet_list[0]):
            sheet_list[0][col_num] = sheet_list[0][col_num].strip()
            if sheet_list[0][col_num] == "":
                for row_num in range( len( sheet_list ) ):
                    del sheet_list[row_num][col_num]
            else:
                col_num += 1
        # Заполнение ячеек в столбцах данных, содержащих ссылку 
        # "^" на расположенную выше ячейку:
        for col_num in range( len( sheet_list[0] ) ):
            startData = ""
            for row_num in range(1, len(sheet_list)):
                data = str(sheet_list[row_num][col_num])
                if data != "^":
                    startData = data
                sheet_list[row_num][col_num] = startData # замена текущих данных на обработанные
        # book_dict[sheetname] = sheet_list

def to_dict_dict_list(dict_list_list):
    """ Для удобства дальнейшей обработки 
        ИЗ словаря двух вложенных списков 
        ДЕЛАЕМ словарь словарей списков 
    """
    dict_dict_list = dict_list_list.copy()
    for sheetname in list(dict_dict_list.keys()):
        sheet_dict = {}
        for col_num in range(len(dict_dict_list[sheetname][0])):
            column = []
            for row_num in range(1, len(dict_dict_list[sheetname])):
                column.append(dict_dict_list[sheetname][row_num][col_num].strip())
            sheet_dict[dict_dict_list[sheetname][0][col_num]] = column
        dict_dict_list[sheetname] = sheet_dict
    return dict_dict_list
def to_dict_list_list(dict_dict_list):
    """ Для экспорта
        ИЗ словаря словарей списков
        ДЕЛАЕМ словарь двух вложенных списков
    """
    dict_list_list = dict_dict_list.copy()
    for sheetname in list(dict_list_list.keys()):
        sheet_list = []
        sheet_list.append(list(dict_list_list[sheetname].keys()))
        for row_num in range(len(dict_list_list[sheetname][sheet_list[0][0]])):
            row = []
            for key in sheet_list[0]:
                row.append(dict_list_list[sheetname][key][row_num])
            sheet_list.append(row)
        dict_list_list.update({sheetname : sheet_list})
    return dict_list_list

def counter(start=0, step=1, number=1):
    """ Возвращает функцию-счётчик, которая, в свою очередь, 
        возвращает значение счётчика, начиная со start, с шагом step, 
        приращение значения происходит при количестве number вызовов функции-счётчика
    """
    i = 0 # Количество вызовов с последнего сброса (но это не точно)
    count = start # Переменная счётчика
    def incrementer():
        nonlocal i, count, step, number
        i += 1
        if i > number:
            i = 1
            count += step
        return count
    return incrementer

# Кабели & жилы
def stage0(mark_data_DDL):
    """ Удаление строк с текстом "не печатать" в столбце Печать(кабели) \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Печать(кабели) + КАБЕЛЬ(кабели) -> все(кабели) + все(жилы)
    """
    cable_nums = []
    for index, value in enumerate(mark_data_DDL["кабели"]["Печать"]):
        if value == "не печатать":
            cable_nums.append(mark_data_DDL["кабели"]["КАБЕЛЬ"][index])
    i = 0
    while i < len(mark_data_DDL["кабели"]["КАБЕЛЬ"]):
        # Если столбец содержит номер(название) кабеля из списка на удаление cable_nums, то...
        if mark_data_DDL["кабели"]["КАБЕЛЬ"][i] in cable_nums:
            # ...удаляем строку с информацией об этом кабеле
            for key in list(mark_data_DDL["кабели"].keys()):
                del mark_data_DDL["кабели"][key][i]
        else:
            i += 1
    i = 0
    while i < len(mark_data_DDL["жилы"]["Кабель"]):
        if mark_data_DDL["жилы"]["Кабель"][i] in cable_nums:
            for key in list(mark_data_DDL["жилы"].keys()):
                del mark_data_DDL["жилы"][key][i]
        else:
            i += 1
def stage1(mark_data_DDL):
    """ Формирование столбцов Жил, Сечение, Занято из столбца Структура \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Структура(кабели) -> Жил(кабели) + Сечение(кабели) + Занято(кабели)
    """
    mark_data_DDL["кабели"]["Жил"] = []
    mark_data_DDL["кабели"]["Сечение"] = []
    mark_data_DDL["кабели"]["Занято"] = []
    for _, value in enumerate(mark_data_DDL["кабели"]["Структура"]):
        if value in minus:
            value = ""
        splitslash = re.split(r"[/\\]", value)
        splitx = re.split(r"[×xXхХ*]", splitslash[0])
        str = ""
        for ind, val in enumerate(splitx):
            if ind < len(splitx)-2:
                str += val.strip()+"×"
            if ind == len(splitx)-2:
                str += val.strip()
        mark_data_DDL["кабели"]["Жил"].append(str)
        mark_data_DDL["кабели"]["Сечение"].append(float(splitx[-1].replace(",",".")) if splitx[-1] != "" else "")
        mark_data_DDL["кабели"]["Занято"].append(int(splitslash[-1]) if splitslash[-1] != "" else "")
def stage2(mark_data_DDL):
    """ Добавление строк текста для маркировки одной из резервных жил кабеля \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Занято(кабели) + Жил(кабели) -> ЖИЛА(жилы)
    """
    cable_nums = []
    for index, value in enumerate(mark_data_DDL["кабели"]["Занято"]):
        if value != "":
            total = 1
            conds = ""
            for j in re.split(r"[×xXхХ*]", mark_data_DDL["кабели"]["Жил"][index]):
                j = j.strip()
                total = total * int(j)
                conds = conds + "×" + j if conds != "" else j
            mark_data_DDL["кабели"]["Жил"][index] = conds
            used = int(value)
            if total > used:
                cable_nums.append(mark_data_DDL["кабели"]["КАБЕЛЬ"][index])
            elif total == used:
                pass
            else:
                raise Exception("Ошибка формата файла")
    i = 1
    prev_cable_num = mark_data_DDL["жилы"]["Кабель"][0]
    while i < len(mark_data_DDL["жилы"]["Кабель"]):
        if prev_cable_num != mark_data_DDL["жилы"]["Кабель"][i]:
            if prev_cable_num in cable_nums:
                for key in list(mark_data_DDL["жилы"].keys()):
                    text = prev_cable_num if key == "ЖИЛА" else ""
                    mark_data_DDL["жилы"][key] = mark_data_DDL["жилы"][key][:i] + [text] + mark_data_DDL["жилы"][key][i:]
                i += 1
            prev_cable_num = mark_data_DDL["жилы"]["Кабель"][i]
        if i == len(mark_data_DDL["жилы"]["Кабель"])-1 and mark_data_DDL["жилы"]["Кабель"][i] in cable_nums:
            for key in list(mark_data_DDL["жилы"].keys()):
                text = mark_data_DDL["жилы"]["Кабель"][i] if key == "ЖИЛА" else ""
                mark_data_DDL["жилы"][key].append(text)
            i += 1
        i += 1
def stage3(mark_data_DDL):
    """ Формирование столбца Сечение(жилы) \n
        Обработка: источник -> приёмник \n
        Столбец(лист): КАБЕЛЬ(кабели) + Сечение(кабели) + Кабель(жилы) -> Сечение(жилы)
    """
    croSSect_dict = {}
    mark_data_DDL["жилы"]["Сечение"] = [0.0 for _ in range(len(mark_data_DDL["жилы"]["Кабель"]))]
    for index, value in enumerate(mark_data_DDL["кабели"]["КАБЕЛЬ"]):
        if mark_data_DDL["кабели"]["Сечение"][index] != "":
            croSSect_dict[value] = float(mark_data_DDL["кабели"]["Сечение"][index])
    for index, value in enumerate(mark_data_DDL["жилы"]["Кабель"]):
        cable_num = mark_data_DDL["жилы"]["ЖИЛА"][index] if value == "" else value
        if cable_num in list(croSSect_dict.keys()):
            mark_data_DDL["жилы"]["Сечение"][index] = croSSect_dict[cable_num]
def stage4(mark_data_DDL):
    """ Формирование столбца ЖилСечение(кабели) \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Жил(кабели) + Сечение(кабели) -> ЖилСечение(кабели)
    """
    mark_data_DDL["кабели"]["ЖилСечение"] = ["" for _ in range(len(mark_data_DDL["кабели"]["Жил"]))]
    for index, value in enumerate(mark_data_DDL["кабели"]["Жил"]):
        if value != "" and mark_data_DDL["кабели"]["Сечение"][index] != "":
            sep = "×"
        else:
            sep = ""
        mark_data_DDL["кабели"]["ЖилСечение"][index] = value + sep + str(mark_data_DDL["кабели"]["Сечение"][index]).replace(".", ",")
def stage5(mark_data_DDL):
    """ Обработка столбца Длина(кабели) \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Длина(кабели) -> Длина(кабели)
    """
    for index, value in enumerate(mark_data_DDL["кабели"]["Длина"]):
        if value != "":
            mark_data_DDL["кабели"]["Длина"][index] = "L = {0} м".format(value.replace(".", ","))
def stage6(mark_data_DDL):
    """ Дублирование строк в соответствии с информацией в столбце Кол.(кабели) \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Кол.(кабели) -> все(кабели)
    """
    i = 0
    while i < len(mark_data_DDL["кабели"]["Кол."]):
        num_cop = int(mark_data_DDL["кабели"]["Кол."][i]) if mark_data_DDL["кабели"]["Кол."][i] != "" else 0
        if num_cop > 0:
            for _ in range(num_cop - 1):
                for key in list(mark_data_DDL["кабели"].keys()):
                    if i == len(mark_data_DDL["кабели"]["Кол."]):
                        mark_data_DDL["кабели"][key].append(mark_data_DDL["кабели"][key][i])
                    else:
                        # mark_data_DDL["кабели"][key] = mark_data_DDL["кабели"][key][:i+1] + [mark_data_DDL["кабели"][key][i]] + mark_data_DDL["кабели"][key][i+1:]
                        mark_data_DDL["кабели"][key].insert(i+1, mark_data_DDL["кабели"][key][i])
                i += 1
        else:
            for key in list(mark_data_DDL["кабели"].keys()):
                del mark_data_DDL["кабели"][key][i]
            i -= 1
        i += 1
def stage7(mark_data_DDL):
    """ Сортировка: сначала маркировка начала всех жил, затем маркировка конца всех жил
        Обработка: источник -> приёмник \n
        Столбец(лист): Кабель(жилы) + ЖИЛА(жилы) + Начало(жилы) + Сечение(жилы) \n
                       Кабель(жилы) + ЖИЛА(жилы) + Конец (жилы) + Сечение(жилы) \n
                    -> Кабель(жилы) + ЖИЛА(жилы) + Адрес (жилы) + Сечение(жилы) \n
    """
    for index, _ in enumerate(mark_data_DDL["жилы"]["Начало"]):
        for value in ("XT:", "XT1:", "XT2:", "XT3:", "XT4:"):
            if mark_data_DDL["жилы"]["Начало"][index].find(value) != -1:
                mark_data_DDL["жилы"]["Начало"][index], mark_data_DDL["жилы"]["Конец"][index] = mark_data_DDL["жилы"]["Конец"][index], mark_data_DDL["жилы"]["Начало"][index]
                continue
    mark_data_DDL["жилы"]["Кабель"] = mark_data_DDL["жилы"]["Кабель"] + ["← Нач"] + mark_data_DDL["жилы"]["Кабель"] + [r"КОН"]
    mark_data_DDL["жилы"]["ЖИЛА"] = mark_data_DDL["жилы"]["ЖИЛА"] + [" / "] + mark_data_DDL["жилы"]["ЖИЛА"] + [r":–)"]
    mark_data_DDL["жилы"]["Адрес"] = mark_data_DDL["жилы"]["Начало"] + ["Кон →"] + mark_data_DDL["жилы"]["Конец"] + [r"ЕЦ!"]
    del mark_data_DDL["жилы"]["Начало"]
    del mark_data_DDL["жилы"]["Конец"]
    mark_data_DDL["жилы"]["Сечение"] = mark_data_DDL["жилы"]["Сечение"] + [mark_data_DDL["жилы"]["Сечение"][-1]] + mark_data_DDL["жилы"]["Сечение"] + [mark_data_DDL["жилы"]["Сечение"][-1]]
# Элементы
def stage8(mark_data_DDL):
    """ Формирование листов "элементЗПО" и "элементНПО" \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Текст(клеммы) + Вид(клеммы) -> Текст(элементЗПО) + Текст(элементНПО)
    """
    # Удалил столбцы Вид1 и Вид2 в файле данных маркировки, добавляю их здесь, чтобы ничего не сломалось
    mark_data_DDL["клеммы"]["Вид1"] = []
    mark_data_DDL["клеммы"]["Вид2"] = []
    for index, value in enumerate(mark_data_DDL["клеммы"]["Текст1"]):
        mark_data_DDL["клеммы"]["Вид1"].append("ЗПО")
        mark_data_DDL["клеммы"]["Вид2"].append("ЗПО")

    def fillSyncGroup(TextZPO, TextNPO):
        """ Перечисление с периодом period значений в столбцах,
            на сколько помню
        """
        while len(Text1) > period / 2:
            if mark_data_DDL["клеммы"]["Вид1"][index-1] == "ЗПО":
                sequencenum = len(TextZPO) // period % maxseqnum
                if   sequencenum == 0:
                    TextZPO.append(Text1.pop(0))
                elif sequencenum == 1:
                    TextZPO.append(Text2.pop(0))
            if mark_data_DDL["клеммы"]["Вид1"][index-1] == "НПО":
                sequencenum = len(TextNPO) // period % maxseqnum
                if   sequencenum == 0:
                    TextNPO.append(Text1.pop(0))
                elif sequencenum == 1:
                    TextNPO.append(Text2.pop(0))
        if mark_data_DDL["клеммы"]["Вид1"][index-1] == "ЗПО":
            TextZPO += Text1 + Text2
        if mark_data_DDL["клеммы"]["Вид1"][index-1] == "НПО":
            TextNPO += Text1 + Text2
        Text1.clear()
        Text2.clear()
    period = 10
    maxseqnum = 2
    Text1 = []
    Text2 = []
    TextZPO = []
    TextNPO = []
    for index, value in enumerate(mark_data_DDL["клеммы"]["Вид1"]):
        if value == mark_data_DDL["клеммы"]["Вид2"][index]:
            flagSyncGroup = True
            Text1.append(mark_data_DDL["клеммы"]["Текст1"][index])
            Text2.append(mark_data_DDL["клеммы"]["Текст2"][index])
        else:
            if flagSyncGroup:
                flagSyncGroup = False
                fillSyncGroup(TextZPO, TextNPO)
            if   mark_data_DDL["клеммы"]["Вид1"][index] == "ЗПО":
                TextZPO.append(mark_data_DDL["клеммы"]["Текст1"][index])
            elif mark_data_DDL["клеммы"]["Вид1"][index] == "НПО":
                TextNPO.append(mark_data_DDL["клеммы"]["Текст1"][index])
            if   mark_data_DDL["клеммы"]["Вид2"][index] == "ЗПО":
                TextZPO.append(mark_data_DDL["клеммы"]["Текст2"][index])
            elif mark_data_DDL["клеммы"]["Вид2"][index] == "НПО":
                TextNPO.append(mark_data_DDL["клеммы"]["Текст2"][index])
    if flagSyncGroup:
        flagSyncGroup = False
        fillSyncGroup(TextZPO, TextNPO)
    mark_data_DDL["элементЗПО"] = {}
    mark_data_DDL["элементНПО"] = {}
    mark_data_DDL["элементЗПО"]["Текст"] = TextZPO
    mark_data_DDL["элементНПО"]["Текст"] = TextNPO



    # for i in range(len(mark_data_DDL["клеммы"]["Вид1"])):
    #     if mark_data_DDL["клеммы"]["Вид1"][i] == "НПО":
    #         mark_data_DDL["элементНПО"]["Текст"].append(mark_data_DDL["клеммы"]["Текст1"][i])
    #     if mark_data_DDL["клеммы"]["Вид1"][i] == "ЗПО":
    #         mark_data_DDL["элементЗПО"]["Текст"].append(mark_data_DDL["клеммы"]["Текст1"][i])
    #     if mark_data_DDL["клеммы"]["Вид2"][i] == "НПО":
    #         mark_data_DDL["элементНПО"]["Текст"].append(mark_data_DDL["клеммы"]["Текст2"][i])
    #     if mark_data_DDL["клеммы"]["Вид2"][i] == "ЗПО":
    #         mark_data_DDL["элементЗПО"]["Текст"].append(mark_data_DDL["клеммы"]["Текст2"][i])
    del mark_data_DDL["клеммы"]
# Проводники
def stage9(mark_data_DDL):
    """ Формирование листа "проводники" \n
        Обработка: источник -> приёмник \n
        Столбец(лист): Все(провод) -> Адрес(провод) + Сечение(провод)
    """
    # Создание списка уникальных названий групп:
    unique_groups = []
    for value in mark_data_DDL["проводники"]["Группа"]:
        if value not in unique_groups:
            unique_groups.append(value)
    # Создание словаря с группами пустых списков:
    groups_DDL = {}
    for group in unique_groups:
        groups_DDL[group] = {}
        for key in list(mark_data_DDL["проводники"].keys()):
            groups_DDL[group][key] = []
    # Заполнение групп пустых списков элементами:
    for index, value in enumerate(mark_data_DDL["проводники"]["Группа"]):
        for key in list(mark_data_DDL["проводники"].keys()):
            groups_DDL[value][key].append(mark_data_DDL["проводники"][key][index])

    # Дублирование записей в каждой группе:
    for group in list(groups_DDL.keys()):
        if int(groups_DDL[group]["Кол."][0]) == 0:
            del groups_DDL[group]
        else:
            for key in list(groups_DDL[group].keys()):
                groups_DDL[group][key] = groups_DDL[group][key] * int(groups_DDL[group]["Кол."][0])
    # Удаление групп, которые содержат значения "не печатать" в столбце "Печать" во ВСЕХ ячейках группы
    for group in list(groups_DDL.keys()):
        notPrintGroup = True
        for item in groups_DDL[group]["Печать"]:
            if item != "не печатать":
                notPrintGroup = False
        if notPrintGroup == True:
            del groups_DDL[group]
    # Инициализация счётчиков для нумерации:
    countK = counter(start=1, number=4)
    countXT1 = counter(start=1)
    countXT2 = counter(start=1)
    countXT3 = counter(start=1)
    countXT4 = counter(start=1)
    separator = " / " # разделитель адресов начала и конца проводника
    # Результирующий словарь:
    conductors = {}
    conductors["single"] = {"Адрес" : [], "Сечение" : [], "Тип" : []}
    conductors["dual"] = {"Адрес" : [], "Сечение" : [], "Тип" : []}
    for group in list(groups_DDL.keys()):
        groups_DDL[group]["Тип"] = []
        for index, _ in enumerate(groups_DDL[group]["Сечение"]):
            # Нумерация элементов K, XT1, XT2, XT3, XT4:
            for key in ("Начало", "Конец"):
                if "K#:" in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = groups_DDL[group][key][index].replace("K#:", "K{0}:".format(countK()))
                if "XT1:#" in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = groups_DDL[group][key][index].replace("XT1:#", "XT1:{0}".format(countXT1()))
                if "XT2:#" in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = groups_DDL[group][key][index].replace("XT2:#", "XT2:{0}".format(countXT2()))
                if "XT3:#" in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = groups_DDL[group][key][index].replace("XT3:#", "XT3:{0}".format(countXT3()))
                if "XT4:#" in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = groups_DDL[group][key][index].replace("XT4:#", "XT4:{0}".format(countXT4()))
            if groups_DDL[group]["Начало"][index] == "":
                MarkType = "single" # Тип маркировки - одиночная. "single" - маркировка проводника на данном конце только адресом подключения этого конца
                groups_DDL[group]["Начало"][index], groups_DDL[group]["Конец"][index] = groups_DDL[group]["Конец"][index], groups_DDL[group]["Начало"][index]
            elif groups_DDL[group]["Конец"][index] == "":
                MarkType = "single"
            else:
                MarkType = "dual" # Тип маркировки - двойная. "dual" - маркировка проводника на данном конце адресом его подключения и адресом подключения противоположного конца этого проводника
                # Формирование текста обратной маркировки:
                begining = groups_DDL[group]["Начало"][index]
                end = groups_DDL[group]["Конец"][index]
                groups_DDL[group]["Начало"][index] = begining + separator + end
                groups_DDL[group]["Конец"][index] = end + separator + begining
            groups_DDL[group]["Тип"].append(MarkType)
        # Удаление строк с меткой "не печатать":
        i = 0
        while i < len(groups_DDL[group]["Печать"]):
            if groups_DDL[group]["Печать"][i] == "не печатать":
                for key in list(groups_DDL[group].keys()):
                    del groups_DDL[group][key][i]
            else:
                i += 1
        # Формирование результирующего словаря:
        crs_sep = groups_DDL[group]["Сечение"][0]
        typ_sep = groups_DDL[group]["Тип"][0]
        if MarkType == "dual":
            conductors[MarkType]["Адрес"] = conductors[MarkType]["Адрес"] + ["Группа {0} →".format(group)] + groups_DDL[group]["Начало"] + ["← Начала / Концы →"] + groups_DDL[group]["Конец"]
            conductors[MarkType]["Сечение"] = conductors[MarkType]["Сечение"] + [crs_sep] + groups_DDL[group]["Сечение"] + [crs_sep] + groups_DDL[group]["Сечение"]
            conductors[MarkType]["Тип"] = conductors[MarkType]["Тип"] + [typ_sep] + groups_DDL[group]["Тип"] + [typ_sep] + groups_DDL[group]["Тип"]
        else:
            conductors[MarkType]["Адрес"] = conductors[MarkType]["Адрес"] + ["Группа {0} →".format(group)] + groups_DDL[group]["Начало"]
            conductors[MarkType]["Сечение"] = conductors[MarkType]["Сечение"] + [crs_sep] + groups_DDL[group]["Сечение"]
            conductors[MarkType]["Тип"] = conductors[MarkType]["Тип"] + [typ_sep] + groups_DDL[group]["Тип"]
    for MarkType in conductors.keys():
        if conductors[MarkType]["Адрес"] != []:
            conductors[MarkType]["Адрес"].append(r"КОНЕЦ! :–)")
            conductors[MarkType]["Сечение"].append(crs_sep)
            conductors[MarkType]["Тип"].append(MarkType)
    mark_data_DDL["проводники"] = {"Адрес" : conductors["dual"]["Адрес"] +   conductors["single"]["Адрес"], \
                                 "Сечение" : conductors["dual"]["Сечение"] + conductors["single"]["Сечение"], \
                                     "Тип" : conductors["dual"]["Тип"] +     conductors["single"]["Тип"]}

def convert_to_transfer(book_dict):
    """ Создание словаря согласно формату трансферного файла
    """
    book_d = OrderedDict() # словарь для формирования трансферного файла
    for newsheetname in newsheetnames:
        sheetname = sheet.column[newsheetname][0][1:-1]
        sheet_dict = OrderedDict()
        for head in sheet.column[newsheetname][1:]: # формирование заголовков
            if head == "":
                break
            if sheetname in book_dict:
                sheet_dict.update({head : book_dict[sheetname][head]})
            else:
                sheet_dict.update({head : []})
        book_d.update({newsheetname: sheet_dict})
    return book_d

# __________________________________________________________

def proc_mark_file(src, dst):
    try:
        mark_data_DLL = pyexcel.get_book_dict(file_name=src) # ИМПОРТ ИЗ ФАЙЛА ДАННЫХ МАРКИРОВКИ
    except:
        return "Ошибка открытия файла данных"
    # try:
    preproc(mark_data_DLL) # Предобработка словаря матриц
    # except:
        # return "Ошибка предобработки"
    try:
        # Преобразование в словарь словарей списков для дальнейшей обработки
        mark_data_DDL = to_dict_dict_list(mark_data_DLL)
    except:
        return "Ошибка преобразования для обработки"
    # Обработка
    if "кабели" in mark_data_DDL and "жилы" in mark_data_DDL:
        try:
            stage0(mark_data_DDL)
        except:
            return "Ошибка формата. stage0. Столбец(лист): Печать(кабели) + КАБЕЛЬ(кабели) -> все(кабели) + все(жилы)"
        try:
            stage1(mark_data_DDL) # Структура -> кабели
        except:
            return "Ошибка формата. stage1. Столбец(лист): Структура(кабели) -> Жил(кабели) + Сечение(кабели) + Занято(кабели)"
        try:
            stage2(mark_data_DDL) # Занято + Жил -> жилы
        except:
            return "Ошибка формата. stage2. Столбец(лист): Занято(кабели) + Жил(кабели) -> ЖИЛА(жилы)"
        try:
            stage3(mark_data_DDL) # КАБЕЛЬ + Сечение -> жилы
        except:
            return "Ошибка формата. stage3. Столбец(лист): КАБЕЛЬ(кабели) + Сечение(кабели) + Кабель(жилы) -> Сечение(жилы)"
        try:
            stage4(mark_data_DDL) # ЖилСечение
        except:
            return "Ошибка формата. stage4. Столбец(лист): Жил(кабели) + Сечение(кабели) -> ЖилСечение(кабели)"
        try:
            stage5(mark_data_DDL) # Длина
        except:
            return "Ошибка формата. stage5. Столбец(лист): Длина(кабели) -> Длина(кабели)"
        try:
            stage6(mark_data_DDL) # Кол.
        except:
            return "Ошибка формата. stage6. Столбец(лист): Кол.(кабели) -> все(кабели)"
        try:
            stage7(mark_data_DDL) # Кабель + ЖИЛА + Начало + Сечение   +   # Кабель + ЖИЛА + Конец  + Сечение
        except:
            return "Ошибка формата. stage7. Столбец(лист): Кабель(жилы) + ЖИЛА(жилы) + Начало(жилы) + Сечение(жилы)"
    if "клеммы" in mark_data_DDL:
        try:
            stage8(mark_data_DDL)
        except:
            return "Ошибка формата. stage8. Столбец(лист): Текст(клеммы) + Вид(клеммы) -> Текст(элементЗПО) + Текст(элементНПО)"
    if "проводники" in mark_data_DDL:
        # try:
        stage9(mark_data_DDL)
        # except:
        #     return "Ошибка формата. stage9. Столбец(лист): Все(провод) -> Адрес(провод) + Сечение(провод)"
    try:
        # Преобразование в соответствии с форматом трансферного файла
        transfer_DDL = convert_to_transfer(mark_data_DDL)
    except:
        return "Ошибка преобразования к формату трансферного файла"
    try:
        # Обратное преобразование для экспорта
        transfer_DLL = to_dict_list_list(transfer_DDL)
    except:
        return "Ошибка преобразования для сохранения"
    try:
        pyexcel.save_book_as(bookdict=transfer_DLL, dest_file_name=dst) # ЭКСПОРТ В ТРАНСФЕРНЫЙ ФАЙЛ
    except:
        return "Ошибка сохранения трансферного файла"
    return 0

def GUImain():
    """ Главное окно программы
    """
    LayoutRange = 20
    LayoutDistance = 20
    MainLayout = \
    [
        [
            sg.Text("Производитель принтера:", size=(max(len("Файл данных:"), len("Производитель принтера:")), 1), pad=((0,0),(10,LayoutRange))),
            sg.DropDown(printprograms, size=(max([len(i) for i in printprograms])+3,1), key="#ProgramSelection", enable_events=True, pad=((0,LayoutDistance),(10,LayoutRange))),
            sg.Button("Установить пакет поддержки", size=(len("Установить пакет поддержки"),1), key="#InstallPack", pad=((0,LayoutDistance+128),(10,LayoutRange))),
            sg.Button("Шрифт InconsolataCyr.ttf", size=(len("Шрифт InconsolataCyr.ttf")+3,1), key="#InstallFont", pad=((0,0),(10,LayoutRange)))
        ],
        [
            sg.Text("Файл данных:", size=(max(len("Файл данных:"), len("Производитель принтера:")), 1), pad=((0,0),(0,LayoutRange))),
            sg.InputText("", size=(93,1), key="#FilePath", pad=((0,LayoutDistance+2),(0,LayoutRange))),
            sg.FileBrowse("Открыть", size=(len("Открыть")+1,1), key="#OpenFile", target="#FilePath", enable_events=True, pad=((0,0),(0,LayoutRange)))
        ],
        [
            sg.Button("Обработать", size=(len("Обработать")+5,2), key="#ProcFile", pad=((0,33),(0,10))),
            sg.Button("Импорт", size=(len("Импорт")+5,2), key="#Import", pad=((0,461),(0,10))),
            sg.Button("Справка", size=(len("Справка")+5,2), key="#Man", pad=((0,33),(0,10))),
            sg.Button("Выход", size=(len("Выход")+5,2), key="#Exit", pad=((0,0),(0,10)))
        ]
    ]
    window = sg.Window("{0}   ver. {1}".format(__progName__, __version__), MainLayout)
    while True:
        event, values = window.read()
        print(event, values)

        if   event in (None, "#Exit"):
            break

        elif event == "#ProgramSelection":
            if not proginstalled(values["#ProgramSelection"]):
                flagproginstalled = False
                sg.PopupError("Программное обеспечение для печати не установлено", title="Ошибка!", keep_on_top=True)
            else:
                flagproginstalled = True
                importfilepath = os.path.join(data("Программы", "KEY", values["#ProgramSelection"])["datapath"], importfilename)
                if not packinstalled(values["#ProgramSelection"]):
                    flagpackinstalled = False
                    sg.PopupError("Пакет поддержки программы печати не установлен или повреждён. Для установки нажмите кнопку \"Установить пакет поддержки\"", title="Ошибка!", keep_on_top=True)
                else:
                    flagpackinstalled = True

        elif event == "#InstallPack":
            if values["#ProgramSelection"] == "":
                sg.Popup("Выберите производителя принтера", title="Ошибка!", keep_on_top=True)
            elif flagproginstalled == False:
                sg.PopupError("Программное обеспечение для печати не установлено", title="Ошибка!", keep_on_top=True)
            else:
                if installpack(values["#ProgramSelection"]):
                    sg.Popup("Пакет поддержки программы печати успешно установлен", title="Успех", keep_on_top=True)
                    flagpackinstalled = True
                else:
                    sg.PopupError("Не удалось установить пакет поддержки программы печати", title="Ошибка!", keep_on_top=True)

        elif event == "#InstallFont":
            fontfile = data("Пути", "KEY", "VALUE")["font"]
            if os.path.exists(fontfile):
                os.startfile(fontfile)
            else:
                sg.PopupError("Не найден файл шрифта. Переустановите программу MarkV", title="Ошибка!", keep_on_top=True)

        elif event == "#ProcFile":
            if values["#ProgramSelection"] == "":
                sg.Popup("Выберите производителя принтера", title="Ошибка!", keep_on_top=True)
            elif flagproginstalled == False:
                sg.PopupError("Программное обеспечение для печати не установлено", title="Ошибка!", keep_on_top=True)
            elif values["#FilePath"] == "":
                sg.Popup("Выберите файл данных маркировки", title="Ошибка!", keep_on_top=True)
            elif not os.path.exists(values["#FilePath"]):
                sg.PopupError("Указанный файл данных не найден", title="Ошибка!", keep_on_top=True)
            else:
                # try:
                ret = proc_mark_file(values["#FilePath"], importfilepath)
                # except:
                    # ret = "Неизвестная ошибка формата данных"
                if ret == 0:
                    if sg.PopupYesNo("Файл данных успешно обработан. \nСформирован и сохранён трансферный файл. \n\nОткрыть трансферный файл для просмотра?\n", title="Выполнено!", keep_on_top=True) == "Yes":
                        os.startfile(importfilepath) # Открытие трансферного файла для проверки
                else:
                    sg.Popup(ret, title="Ошибка!", keep_on_top=True)

        elif event == "#Import":
            if values["#ProgramSelection"] == "":
                sg.Popup("Выберите производителя принтера", title="Ошибка!", keep_on_top=True)
            elif flagproginstalled == False:
                sg.PopupError("Программное обеспечение для печати не установлено", title="Ошибка!", keep_on_top=True)
            elif flagpackinstalled == False:
                sg.PopupError("Пакет поддержки программы печати не установлен или повреждён. Для установки нажмите кнопку \"Установить пакет поддержки\"", title="Ошибка!", keep_on_top=True)
            elif not os.path.exists(importfilepath):
                sg.PopupError("Не найден трансферный файл. Чтобы его сформировать выберите файл данных и нажмите кнопку \"Обработать\"", title="Ошибка!", keep_on_top=True)
            else:
                callstring  = "\"{0}\"".format(data("Программы", "KEY", values["#ProgramSelection"])["programpath"])
                paramstring = "\"{0}\"".format(data("Программы", "KEY", values["#ProgramSelection"])["paramCMD"])
                callstring = callstring + " " + paramstring
                print(callstring)
                subprocess.Popen(callstring, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        elif event == "#Man":
            manfile = data("Пути", "KEY", "VALUE")["man"]
            if os.path.exists(manfile):
                os.startfile(manfile)
            else:
                sg.PopupError("Не найден файл справки. Переустановите программу MarkV", title="Ошибка!", keep_on_top=True)
            # subprocess.call("Acrobat.exe /A page=3 {0}".format(manfile))

    window.close()



if __name__ == "__main__":
    GUImain()
