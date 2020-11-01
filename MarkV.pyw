"""
Main module of MarkV programm

Данный скрипт следует непосредственно выполнять в интерпретаторе Python.
Программа не имеет параметров командной строки.
"""

import os
import re
import shutil
import subprocess

import PySimpleGUI as sg
import pyexcel
from pyexcel._compact import OrderedDict



__progName__ = 'MarkV'   # Name of programm
__version__ = '2020_11_02'   # Version of programm

# Themes: 'Default 1', 'DarkTanBlue', 'System Default 1'...
GUI_THEME = 'System Default 1'

# Регистро независимый поиск соответствия:
# Знаки начала имени игнорируемых листов
MINUS = ('-', '−', '–')
# Имена листов, которые обрабатывает программа
WORK_SHEETS = ('клеммы', 'провода', 'кабели', 'жилы')
# Имена листов, которые игнорирует программа
NOT_WORK_SHEETS = ()

# Содержимое пустой ячейки
HOLLOW = ('', None)
# Ссылка на расположенную выше ячейку в файле данных
LINK_UP = '^'

# Разделитель адресов начала и конца проводника в зеркальной маркировке
SEPARATOR = ' / '

PROG_DATA_FILE_NAME = 'prog_data.xlsx'
PROGRAM_DATA = {}


def read_program_data():
    """Загрузка файла внутренних параметров работы программы

    Функция возвращает словарь словарей списков,
    соответсвующий данным в загружаемой книге Excel

    """
    global PROGRAM_DATA
    data_book_path = os.path.normpath(os.path.join(os.path.dirname(__file__),
                                                   PROG_DATA_FILE_NAME))
    data_book = pyexcel.get_book(file_name=data_book_path)
    for data_book_sheet in data_book:
        data_book_sheet.name_columns_by_row(0)
        PROGRAM_DATA[data_book_sheet.name] = data_book_sheet.to_dict()
    return PROGRAM_DATA


def get_programm_data(page_name, key=None, value=None):
    """Получить данные программы

    Возвращает словарь, с ключами и значениями из двух столбцов таблицы
    на листе page_name в PROGRAM_DATA. Заголовки столбцов в переменных:
    key - для ключей, value - для значений.
    Если передан только key или value, то возвратиться список из данных
    в столбце с апереданным заголовком.
    При отсутствии key и value функция вернёт в качестве результата
    collections.OrderedDict с заголовками столбцов в качестве ключей, и
    списками из данных в столбце в качестве значений.

    """
    global PROGRAM_DATA
    sheet_data = PROGRAM_DATA[page_name]
    if key is None and value is None:
        return sheet_data

    for header in sheet_data:
        if header == key:
            keys = sheet_data[header]
        if header == value:
            values = sheet_data[header]
    if key is None:
        return values
    if value is None:
        return keys
    dict_ = {}
    for index, key_ in enumerate(keys):
        dict_[key_] = values[index]
    return dict_


def prog_installed_check(current_print_program):
    """Возвращает True, если программа для печати установлена и
    False, если нет или не найдена.

    """
    return os.path.exists(gpd('Программы',
                              'KEY',
                              current_print_program)['program_path']) and \
           os.path.exists(gpd('Программы',
                              'KEY',
                              current_print_program)['pack_path_dst'])


def pack_installed_check(current_print_program):
    """Возвращает True, если пакет поддержки программы установлен и
    False, если не установлен или повреждён.

    """
    src = gpd('Программы', 'KEY', current_print_program)['pack_path_src']
    dst = gpd('Программы', 'KEY', current_print_program)['pack_path_dst']
    for src_tuple in os.walk(src):
        for src_dir in src_tuple[1]:
            path = os.path.join(src_tuple[0], src_dir).replace(src, dst)
            # print(path)
            if not os.path.exists(path):
                return False
        for src_file in src_tuple[2]:
            path = os.path.join(src_tuple[0], src_file).replace(src, dst)
            # print(path)
            if not os.path.exists(path):
                return False
    return True


def install_pack(current_print_program):
    """Устанавливает пакет поддрежки соответствующей программы печати"""
    src = gpd('Программы', 'KEY', current_print_program)['pack_path_src']
    dst = gpd('Программы', 'KEY', current_print_program)['pack_path_dst']
    for src_tuple in os.walk(src):
        for src_dir in src_tuple[1]:
            path = os.path.join(src_tuple[0], src_dir).replace(src, dst)
            if not os.path.exists(path):
                # print(path)
                try:
                    os.mkdir(path)
                except Exception:
                    return False
        for src_file in src_tuple[2]:
            src_file_path = os.path.join(src_tuple[0], src_file)
            dst_file_path = src_file_path.replace(src, dst)
            if not os.path.exists(dst_file_path):
                print('Отсутствует файл:', dst_file_path)
            try:
                shutil.copyfile(src_file_path, dst_file_path)
            except Exception:
                return False
            print('Скопирован файл:', src_file_path, '\n',
                  'в файл:', dst_file_path)
    return True


def preproc(book_dict):
    """Предобработка таблиц данных.

    Удаление:
     - листов, начинающихся со знаков MINUS,
     - листов, отсутствующих в WORK_SHEETS,
     - листов, присутствующих в NOT_WORK_SHEETS,
     - пустых строк,
     - неозаглавленных столбцов.
    Заполнение ячеек, содержащих ссылку '^'

    """
    sheet_names = set(sheet_name for sheet_name in book_dict)
    for sheet_name in sheet_names:
        # Удаление листов с именами, начинающимися с минуса
        if sheet_name.startswith(MINUS):
            del book_dict[sheet_name]
            continue
        # Удаление листов, отсутствующих в WORK_SHEETS
        if WORK_SHEETS and sheet_name not in WORK_SHEETS:
            del book_dict[sheet_name]
            continue
        # Удаление листов, присутствующих в NOT_WORK_SHEETS
        if NOT_WORK_SHEETS and sheet_name in NOT_WORK_SHEETS:
            del book_dict[sheet_name]
            continue
    # Удаление пробелов вначале и конце строки в ячейках
    for sheet_name in book_dict:
        for row_num in range(len(book_dict[sheet_name])):
            for col_num in range(len(book_dict[sheet_name][0])):
                if isinstance(book_dict[sheet_name][row_num][col_num],
                              str):
                    book_dict[sheet_name][row_num][col_num] = \
                    book_dict[sheet_name][row_num][col_num].strip()
    for sheet_name in book_dict:
        sheet_list = book_dict[sheet_name]
        # Удаление строк содержащих
        # только пустые или заполненные пробелами ячейки
        row_num = 0
        while row_num < len(sheet_list):
            row_empty = True
            for col_num in range(len(sheet_list[0])):
                if sheet_list[row_num][col_num] not in HOLLOW:
                    row_empty = False
                    break
            if row_empty:
                del sheet_list[row_num]
            else:
                row_num += 1
        # Удаление неозаглавленных столбцов данных:
        col_num = 0
        while col_num < len(sheet_list[0]):
            # sheet_list[0][col_num] = sheet_list[0][col_num].strip()
            if sheet_list[0][col_num] in HOLLOW:
                for row in sheet_list:
                    del row[col_num]
            else:
                col_num += 1
        # Заполнение ячеек в столбцах данных,
        # содержащих ссылку LINK_UP ('^')
        # на расположенную выше ячейку:
        for col_num in range(len(sheet_list[0])):
            start_content = LINK_UP
            for row in sheet_list:
                content = row[col_num]
                if content == LINK_UP:
                    row[col_num] = start_content
                else:
                    start_content = content

        # book_dict[sheet_name] = sheet_list


def to_dict_dict_list(dict_list_list):
    """Конвертер структуры dict_list_list в dict_dict_list

    Для удобства дальнейшей обработки
    ИЗ словаря двух вложенных списков
    ДЕЛАЕМ словарь словарей списков

    """
    dict_dict_list = {}
    for sheet_name in dict_list_list:
        sheet_dict = {}
        for col_num in range(len(dict_list_list[sheet_name][0])):
            column = []
            for row in dict_list_list[sheet_name]:
                column.append(row[col_num])
            sheet_dict[dict_list_list[sheet_name][0][col_num]] = column[1:]
        dict_dict_list[sheet_name] = sheet_dict
    return dict_dict_list


def to_dict_list_list(dict_dict_list):
    """Конвертер структуры dict_dict_list в dict_list_list

    Для экспорта
    ИЗ словаря словарей списков
    ДЕЛАЕМ словарь двух вложенных списков

    """
    dict_list_list = dict_dict_list.copy()
    for sheet_name in list(dict_list_list.keys()):
        sheet_list = []
        sheet_list.append(list(dict_list_list[sheet_name].keys()))
        for row_num in range(len(dict_list_list[sheet_name][sheet_list[0][0]])):
            row = []
            for key in sheet_list[0]:
                row.append(dict_list_list[sheet_name][key][row_num])
            sheet_list.append(row)
        dict_list_list.update({sheet_name: sheet_list})
    return dict_list_list


def counter(start=0, step=1, number=1):
    """Возвращает функцию-счётчик.

    Счётчик значение, начиная со start, с шагом step.
    Приращение значения происходит при
    количестве number вызовов функции-счётчика.

    """
    i = 0  # Количество вызовов с последнего сброса (!УТОЧНИТЬ!)
    count = start  # Переменная счётчика
    def incrementer():
        nonlocal i, count, step, number
        i += 1
        if i > number:
            i = 1
            count += step
        return count
    return incrementer


# Кабели & жилы
def stage0(mark_data_ddl):
    """Удаление строк с текстом 'не печатать'
    в столбце Печать(кабели) \n

    Обработка: источник -> приёмник \n
    Столбец(лист): Печать(кабели) + КАБЕЛЬ(кабели) ->
    все(кабели) + все(жилы)

    """
    cable_nums = []
    for index, value in enumerate(mark_data_ddl['кабели']['Печать']):
        if value == 'не печатать':
            cable_nums.append(mark_data_ddl['кабели']['КАБЕЛЬ'][index])
    i = 0
    while i < len(mark_data_ddl['кабели']['КАБЕЛЬ']):
        # Если столбец содержит номер(название) кабеля
        # из списка на удаление cable_nums, то...
        if mark_data_ddl['кабели']['КАБЕЛЬ'][i] in cable_nums:
            # ...удаляем строку с информацией об этом кабеле
            for key in list(mark_data_ddl['кабели'].keys()):
                del mark_data_ddl['кабели'][key][i]
        else:
            i += 1
    i = 0
    while i < len(mark_data_ddl['жилы']['Кабель']):
        if mark_data_ddl['жилы']['Кабель'][i] in cable_nums:
            for key in list(mark_data_ddl['жилы'].keys()):
                del mark_data_ddl['жилы'][key][i]
        else:
            i += 1


def stage1(mark_data_ddl):
    """Формирование столбцов Жил, Сечение, Занято
    из столбца Структура

    Обработка: источник -> приёмник
    Столбец(лист): Структура(кабели) ->
    Жил(кабели) + Сечение(кабели) + Занято(кабели)

    """
    mark_data_ddl['кабели']['Жил'] = []
    mark_data_ddl['кабели']['Сечение'] = []
    mark_data_ddl['кабели']['Занято'] = []
    for _, value in enumerate(mark_data_ddl['кабели']['Структура']):
        if value in MINUS:
            value = ''
        splitslash = re.split(r'[/\\]', value)
        splitx = re.split(r'[×xXхХ*]', splitslash[0])
        string = ''
        for ind, val in enumerate(splitx):
            if ind < len(splitx) - 2:
                string += val.strip() + '×'
            if ind == len(splitx) - 2:
                string += val.strip()
        mark_data_ddl['кабели']['Жил'].append(string)
        mark_data_ddl['кабели']['Сечение'].append(
              float(splitx[-1].replace(',', '.'))
              if splitx[-1] != ''
              else '')
        mark_data_ddl['кабели']['Занято'].append(
              int(splitslash[-1])
              if splitslash[-1] != ''
              else '')


def stage2(mark_data_ddl):
    """Добавление строк текста для маркировки одной из резервных жил
    кабеля

    Обработка: источник -> приёмник
    Столбец(лист): Занято(кабели) + Жил(кабели) -> ЖИЛА(жилы)

    """
    cable_nums = []
    for index, value in enumerate(mark_data_ddl['кабели']['Занято']):
        if value != '':
            total = 1
            conds = ''
            for j in re.split(r'[×xXхХ*]',
                              mark_data_ddl['кабели']['Жил'][index]):
                j = j.strip()
                total = total * int(j)
                conds = conds + '×' + j if conds != '' else j
            mark_data_ddl['кабели']['Жил'][index] = conds
            used = int(value)
            if total > used:
                cable_nums.append(mark_data_ddl['кабели']['КАБЕЛЬ'][index])
            elif total == used:
                pass
            else:
                raise Exception('Ошибка формата файла')
    i = 1
    prev_cable_num = mark_data_ddl['жилы']['Кабель'][0]
    while i < len(mark_data_ddl['жилы']['Кабель']):
        if prev_cable_num != mark_data_ddl['жилы']['Кабель'][i]:
            if prev_cable_num in cable_nums:
                for key in list(mark_data_ddl['жилы'].keys()):
                    text = prev_cable_num if key == 'ЖИЛА' else ''
                    mark_data_ddl['жилы'][key] = (
                          mark_data_ddl['жилы'][key][:i] +
                          [text] +
                          mark_data_ddl['жилы'][key][i:])
                i += 1
            prev_cable_num = mark_data_ddl['жилы']['Кабель'][i]
        if (i == (len(mark_data_ddl['жилы']['Кабель']) - 1) and
              mark_data_ddl['жилы']['Кабель'][i] in cable_nums):
            for key in list(mark_data_ddl['жилы'].keys()):
                text = (mark_data_ddl['жилы']['Кабель'][i]
                        if key == 'ЖИЛА'
                        else '')
                mark_data_ddl['жилы'][key].append(text)
            i += 1
        i += 1


def stage3(mark_data_ddl):
    """Формирование столбца Сечение(жилы)

    Обработка: источник -> приёмник
    Столбец(лист):
    КАБЕЛЬ(кабели) + Сечение(кабели) + Кабель(жилы) -> Сечение(жилы)

    """
    cros_sect_dict = {}
    mark_data_ddl['жилы']['Сечение'] = \
    [0.0 for _ in range(len(mark_data_ddl['жилы']['Кабель']))]
    for index, value in enumerate(mark_data_ddl['кабели']['КАБЕЛЬ']):
        if mark_data_ddl['кабели']['Сечение'][index] != '':
            cros_sect_dict[value] = \
            float(mark_data_ddl['кабели']['Сечение'][index])
    for index, value in enumerate(mark_data_ddl['жилы']['Кабель']):
        cable_num = (mark_data_ddl['жилы']['ЖИЛА'][index]
                     if value == ''
                     else value)
        if cable_num in list(cros_sect_dict.keys()):
            mark_data_ddl['жилы']['Сечение'][index] = cros_sect_dict[cable_num]


def stage4(mark_data_ddl):
    """Формирование столбца ЖилСечение(кабели)

    Обработка: источник -> приёмник
    Столбец(лист): Жил(кабели) + Сечение(кабели) -> ЖилСечение(кабели)

    """
    mark_data_ddl['кабели']['ЖилСечение'] = \
    ['' for _ in range(len(mark_data_ddl['кабели']['Жил']))]
    for index, value in enumerate(mark_data_ddl['кабели']['Жил']):
        if value != '' and mark_data_ddl['кабели']['Сечение'][index] != '':
            sep = '×'
        else:
            sep = ''
        mark_data_ddl['кабели']['ЖилСечение'][index] = ''.join([
              value,
              sep,
              str(mark_data_ddl['кабели']['Сечение'][index]).replace('.', ',')
              ])


def stage5(mark_data_ddl):
    """Обработка столбца Длина(кабели)

    Обработка: источник -> приёмник
    Столбец(лист): Длина(кабели) -> Длина(кабели)

    """
    for index, value in enumerate(mark_data_ddl['кабели']['Длина']):
        if value != '':
            mark_data_ddl['кабели']['Длина'][index] = \
            'L = {0} м'.format(str(value))


def stage6(mark_data_ddl):
    """Дублирование строк в соответствии с информацией в столбце Кол.(кабели)

    Обработка: источник -> приёмник
    Столбец(лист): Кол.(кабели) -> все(кабели)

    """
    row_num = 0  # Номер текущей строки
    while row_num < len(mark_data_ddl['кабели']['Кол.']):
        if mark_data_ddl['кабели']['Кол.'][row_num] in HOLLOW:
            num_cop = 0
        else:
            num_cop = mark_data_ddl['кабели']['Кол.'][row_num]
        if num_cop == 0:
            for key in mark_data_ddl['кабели']:
                del mark_data_ddl['кабели'][key][row_num]
            if row_num > 0:
                row_num -= 1
        elif num_cop > 1:
            for _ in range(num_cop - 1):
                for key in mark_data_ddl['кабели']:
                    mark_data_ddl['кабели'][key].insert(
                          row_num,
                          mark_data_ddl['кабели'][key][row_num])
                row_num += 1
        row_num += 1


def stage7(mark_data_ddl):
    """Сортировка: сперва маркировка начала всех жил,
    затем маркировка конца всех жил

    Обработка: источник -> приёмник. Столбец(лист):
        Кабель(жилы) + ЖИЛА(жилы) + Начало(жилы) + Сечение(жилы)
        Кабель(жилы) + ЖИЛА(жилы) + Конец (жилы) + Сечение(жилы) ->
        Кабель(жилы) + ЖИЛА(жилы) + Адрес (жилы) + Сечение(жилы)

    """
    for index, value in enumerate(mark_data_ddl['жилы']['Начало']):
        for template in ('XT:', 'XT1:', 'XT2:', 'XT3:', 'XT4:'):
            if value.find(template) != -1:
                (mark_data_ddl['жилы']['Начало'][index],
                 mark_data_ddl['жилы']['Конец'][index]
                 ) = (
                 mark_data_ddl['жилы']['Конец'][index],
                 mark_data_ddl['жилы']['Начало'][index])
                continue
    mark_data_ddl['жилы']['Кабель'] = \
          mark_data_ddl['жилы']['Кабель'] + \
          ['← Нач'] + \
          mark_data_ddl['жилы']['Кабель'] + \
          [r'КОН']
    mark_data_ddl['жилы']['ЖИЛА'] = \
          mark_data_ddl['жилы']['ЖИЛА'] + \
          [SEPARATOR] + \
          mark_data_ddl['жилы']['ЖИЛА'] + \
          [r':–)']
    mark_data_ddl['жилы']['Адрес'] = \
          mark_data_ddl['жилы']['Начало'] + \
          ['Кон →'] + \
          mark_data_ddl['жилы']['Конец'] + \
          [r'ЕЦ!']
    del mark_data_ddl['жилы']['Начало']
    del mark_data_ddl['жилы']['Конец']

    mark_data_ddl['жилы']['Сечение'] = \
          mark_data_ddl['жилы']['Сечение'] + \
          [mark_data_ddl['жилы']['Сечение'][-1]] + \
          mark_data_ddl['жилы']['Сечение'] + \
          [mark_data_ddl['жилы']['Сечение'][-1]]


# Элементы
def stage8(mark_data_ddl):
    """Формирование листов 'элементЗПО' и 'элементНПО'

    Обработка: источник -> приёмник
    Столбец(лист): Текст(клеммы) + Вид(клеммы) ->
    Текст(элементЗПО) + Текст(элементНПО)

    """
    # Удалил столбцы Вид1 и Вид2 в файле данных маркировки.
    # Добавляю их здесь, чтобы ничего не сломалось.
    mark_data_ddl['клеммы']['Вид1'] = []
    mark_data_ddl['клеммы']['Вид2'] = []
    for index, value in enumerate(mark_data_ddl['клеммы']['Текст1']):
        mark_data_ddl['клеммы']['Вид1'].append('ЗПО')
        mark_data_ddl['клеммы']['Вид2'].append('ЗПО')

    def fill_sync_group(text_ZPO, text_NPO):
        """Перечисление с периодом PERIOD значений в столбцах
        (!УТОЧНИТЬ!)

        """
        while len(text_1) > PERIOD / 2:
            if mark_data_ddl['клеммы']['Вид1'][index-1] == 'ЗПО':
                sequence_num = len(text_ZPO) // PERIOD % MAX_SEQ_NUM
                if   sequence_num == 0:
                    text_ZPO.append(text_1.pop(0))
                elif sequence_num == 1:
                    text_ZPO.append(text_2.pop(0))
            if mark_data_ddl['клеммы']['Вид1'][index-1] == 'НПО':
                sequence_num = len(text_NPO) // PERIOD % MAX_SEQ_NUM
                if   sequence_num == 0:
                    text_NPO.append(text_1.pop(0))
                elif sequence_num == 1:
                    text_NPO.append(text_2.pop(0))
        if mark_data_ddl['клеммы']['Вид1'][index-1] == 'ЗПО':
            text_ZPO += text_1 + text_2
        if mark_data_ddl['клеммы']['Вид1'][index-1] == 'НПО':
            text_NPO += text_1 + text_2
        text_1.clear()
        text_2.clear()

    PERIOD = 10  # длина строки из маркировочных табличек
    MAX_SEQ_NUM = 2  # !ПОЯСНИТЬ!
    text_1 = []
    text_2 = []
    text_ZPO = []
    text_NPO = []
    for index, value in enumerate(mark_data_ddl['клеммы']['Вид1']):
        if value == mark_data_ddl['клеммы']['Вид2'][index]:
            flag_sync_group = True
            text_1.append(mark_data_ddl['клеммы']['Текст1'][index])
            text_2.append(mark_data_ddl['клеммы']['Текст2'][index])
        else:
            if flag_sync_group:
                flag_sync_group = False
                fill_sync_group(text_ZPO, text_NPO)

            if mark_data_ddl['клеммы']['Вид1'][index] == 'ЗПО':
                text_ZPO.append(mark_data_ddl['клеммы']['Текст1'][index])
            elif mark_data_ddl['клеммы']['Вид1'][index] == 'НПО':
                text_NPO.append(mark_data_ddl['клеммы']['Текст1'][index])

            if mark_data_ddl['клеммы']['Вид2'][index] == 'ЗПО':
                text_ZPO.append(mark_data_ddl['клеммы']['Текст2'][index])
            elif mark_data_ddl['клеммы']['Вид2'][index] == 'НПО':
                text_NPO.append(mark_data_ddl['клеммы']['Текст2'][index])
    if flag_sync_group:
        flag_sync_group = False
        fill_sync_group(text_ZPO, text_NPO)
    mark_data_ddl['элементЗПО'] = {}
    mark_data_ddl['элементНПО'] = {}
    mark_data_ddl['элементЗПО']['Текст'] = text_ZPO
    mark_data_ddl['элементНПО']['Текст'] = text_NPO

    # for i in range(len(mark_data_ddl['клеммы']['Вид1'])):
    #     if mark_data_ddl['клеммы']['Вид1'][i] == 'НПО':
    #         mark_data_ddl['элементНПО']['Текст'].append(mark_data_ddl['клеммы']['Текст1'][i])
    #     if mark_data_ddl['клеммы']['Вид1'][i] == 'ЗПО':
    #         mark_data_ddl['элементЗПО']['Текст'].append(mark_data_ddl['клеммы']['Текст1'][i])
    #     if mark_data_ddl['клеммы']['Вид2'][i] == 'НПО':
    #         mark_data_ddl['элементНПО']['Текст'].append(mark_data_ddl['клеммы']['Текст2'][i])
    #     if mark_data_ddl['клеммы']['Вид2'][i] == 'ЗПО':
    #         mark_data_ddl['элементЗПО']['Текст'].append(mark_data_ddl['клеммы']['Текст2'][i])

    del mark_data_ddl['клеммы']


# провода
def stage9(mark_data_ddl):
    """Формирование листа 'провода'

    Обработка: источник -> приёмник
    Столбец(лист): Все(провод) -> Адрес(провод) + Сечение(провод)

    """
    # Создание списка уникальных названий групп:
    unique_groups = []
    for value in mark_data_ddl['провода']['Группа']:
        if value not in unique_groups:
            unique_groups.append(value)
    # Создание словаря с группами пустых списков:
    groups_DDL = {}
    for group in unique_groups:
        groups_DDL[group] = {}
        for key in list(mark_data_ddl['провода'].keys()):
            groups_DDL[group][key] = []
    # Заполнение групп пустых списков элементами:
    for index, value in enumerate(mark_data_ddl['провода']['Группа']):
        for key in list(mark_data_ddl['провода'].keys()):
            groups_DDL[value][key].append(mark_data_ddl['провода'][key][index])

    # Дублирование записей в каждой группе:
    for group in list(groups_DDL.keys()):
        if int(groups_DDL[group]['Кол.'][0]) == 0:
            del groups_DDL[group]
        else:
            for key in list(groups_DDL[group].keys()):
                groups_DDL[group][key] = (groups_DDL[group][key] * int(
                                          groups_DDL[group]['Кол.'][0]))
    # Удаление групп, которые содержат значения 'не печатать'
    # в столбце 'Печать' во ВСЕХ ячейках группы
    for group in list(groups_DDL.keys()):
        not_print_group = True
        for item in groups_DDL[group]['Печать']:
            if item != 'не печатать':
                not_print_group = False
        if not_print_group:
            del groups_DDL[group]
    # Инициализация счётчиков для нумерации:
    count_K = counter(start=1, number=4)
    count_XT1 = counter(start=1)
    count_XT2 = counter(start=1)
    count_XT3 = counter(start=1)
    count_XT4 = counter(start=1)
    # Результирующий словарь:
    conductors = {}
    conductors['single'] = {'Адрес': [], 'Сечение': [], 'Тип': []}
    conductors['mirror'] = {'Адрес': [], 'Сечение': [], 'Тип': []}
    for group in list(groups_DDL.keys()):
        groups_DDL[group]['Тип'] = []
        for index, _ in enumerate(groups_DDL[group]['Сечение']):
            # Нумерация элементов K, XT1, XT2, XT3, XT4:
            for key in ('Начало', 'Конец'):
                if 'K#:' in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = \
                    groups_DDL[group][key][index].replace(
                          'K#:',
                          'K{0}:'.format(count_K()))
                if 'XT1:#' in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = \
                    groups_DDL[group][key][index].replace(
                          'XT1:#',
                          'XT1:{0}'.format(count_XT1()))
                if 'XT2:#' in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = \
                    groups_DDL[group][key][index].replace(
                          'XT2:#',
                          'XT2:{0}'.format(count_XT2()))
                if 'XT3:#' in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = \
                    groups_DDL[group][key][index].replace(
                          'XT3:#',
                          'XT3:{0}'.format(count_XT3()))
                if 'XT4:#' in groups_DDL[group][key][index]:
                    groups_DDL[group][key][index] = \
                    groups_DDL[group][key][index].replace(
                          'XT4:#',
                          'XT4:{0}'.format(count_XT4()))
            if groups_DDL[group]['Начало'][index] == '':
                # Тип маркировки - одиночная.
                # 'single' - маркировка проводника на данном конце
                # только адресом подключения этого конца
                mark_type = 'single'
                (groups_DDL[group]['Начало'][index],
                 groups_DDL[group]['Конец'][index]
                 ) = (
                 groups_DDL[group]['Конец'][index],
                 groups_DDL[group]['Начало'][index])

            elif groups_DDL[group]['Конец'][index] == '':
                mark_type = 'single'
            else:
                # Тип маркировки - зеркальная.
                # 'mirror' - маркировка проводника на данном конце
                # адресом его подключения и адресом подключения
                # противоположного конца этого проводника
                mark_type = 'mirror'
                # Формирование текста обратной маркировки:
                begining = groups_DDL[group]['Начало'][index]
                end = groups_DDL[group]['Конец'][index]
                groups_DDL[group]['Начало'][index] = ''.join([begining,
                                                              SEPARATOR,
                                                              end])
                groups_DDL[group]['Конец'][index] = ''.join([end,
                                                             SEPARATOR,
                                                             begining])
            groups_DDL[group]['Тип'].append(mark_type)
        # Удаление строк с меткой 'не печатать':
        i = 0
        while i < len(groups_DDL[group]['Печать']):
            if groups_DDL[group]['Печать'][i] == 'не печатать':
                for key in list(groups_DDL[group].keys()):
                    del groups_DDL[group][key][i]
            else:
                i += 1
        # Формирование результирующего словаря:
        crs_sep = groups_DDL[group]['Сечение'][0]
        typ_sep = groups_DDL[group]['Тип'][0]
        if mark_type == 'mirror':
            conductors[mark_type]['Адрес'] = \
                  conductors[mark_type]['Адрес'] + \
                  ['Группа {0} →'.format(group)] + \
                  groups_DDL[group]['Начало'] + \
                  ['← Начала{0}Концы →'.format(SEPARATOR)] + \
                  groups_DDL[group]['Конец']
            conductors[mark_type]['Сечение'] = \
                  conductors[mark_type]['Сечение'] + \
                  [crs_sep] + \
                  groups_DDL[group]['Сечение'] + \
                  [crs_sep] + \
                  groups_DDL[group]['Сечение']
            conductors[mark_type]['Тип'] = \
                  conductors[mark_type]['Тип'] + \
                  [typ_sep] + \
                  groups_DDL[group]['Тип'] + \
                  [typ_sep] + \
                  groups_DDL[group]['Тип']
        else:
            conductors[mark_type]['Адрес'] = \
                  conductors[mark_type]['Адрес'] + \
                  ['Группа {0} →'.format(group)] + \
                  groups_DDL[group]['Начало']
            conductors[mark_type]['Сечение'] = \
                  conductors[mark_type]['Сечение'] + \
                  [crs_sep] + \
                  groups_DDL[group]['Сечение']
            conductors[mark_type]['Тип'] = \
                  conductors[mark_type]['Тип'] + \
                  [typ_sep] + \
                  groups_DDL[group]['Тип']
    for mark_type in conductors:
        if conductors[mark_type]['Адрес'] != []:
            conductors[mark_type]['Адрес'].append(r'КОНЕЦ! :–)')
            conductors[mark_type]['Сечение'].append(crs_sep)
            conductors[mark_type]['Тип'].append(mark_type)
    mark_data_ddl['провода'] = {
          'Адрес': conductors['mirror']['Адрес'] + \
                   conductors['single']['Адрес'],
          'Сечение': conductors['mirror']['Сечение'] + \
                     conductors['single']['Сечение'],
          'Тип': conductors['mirror']['Тип'] + \
                 conductors['single']['Тип']
          }


def convert_to_transfer(book_dict):
    """Создание словаря согласно формату трансферного файла"""
    transfer_book = OrderedDict()
    for new_sheet_name in gpd('Трансфер'):
        sheet_name = gpd('Трансфер', new_sheet_name)[0][1:-1]
        transfer_book_sheet = OrderedDict()
        # Формирование заголовков
        for head in gpd('Трансфер', new_sheet_name)[1:]:
            if head == '':
                break
            if sheet_name in book_dict:
                transfer_book_sheet.update({head : book_dict[sheet_name][head]})
            else:
                transfer_book_sheet.update({head : []})
        transfer_book.update({new_sheet_name: transfer_book_sheet})
    return transfer_book



# __________________________________________________________



def proc_mark_file(src, dst):
    """Главная функция обработки файла маркировки

    """
    try:
        # ИМПОРТ ИЗ ФАЙЛА ДАННЫХ МАРКИРОВКИ
        mark_data_dll = pyexcel.get_book_dict(file_name=src)
    except:
        return \
"""Ошибка открытия файла данных"""
    try:
        preproc(mark_data_dll)  # Предобработка словаря матриц
    except:
        return \
"""Ошибка предобработки"""
    try:
        # Преобразование в словарь словарей списков
        # для дальнейшей обработки
        mark_data_ddl = to_dict_dict_list(mark_data_dll)
    except:
        return \
"""Ошибка преобразования для обработки"""
    # Обработка
    if 'кабели' in mark_data_ddl and 'жилы' in mark_data_ddl:
        try:
            stage0(mark_data_ddl)
        except:
            return \
"""Ошибка формата. stage0. Столбец(лист):
Печать(кабели) + КАБЕЛЬ(кабели) ->
все(кабели) + все(жилы)"""
        try:
            stage1(mark_data_ddl)  # Структура -> кабели
        except:
            return \
"""Ошибка формата. stage1. Столбец(лист):
Структура(кабели) ->
Жил(кабели) + Сечение(кабели) + Занято(кабели)"""
        try:
            stage2(mark_data_ddl)  # Занято + Жил -> жилы
        except:
            return \
"""Ошибка формата. stage2. Столбец(лист):
Занято(кабели) + Жил(кабели) ->
ЖИЛА(жилы)"""
        try:
            stage3(mark_data_ddl)  # КАБЕЛЬ + Сечение -> жилы
        except:
            return \
"""Ошибка формата. stage3. Столбец(лист):
КАБЕЛЬ(кабели) + Сечение(кабели) + Кабель(жилы) ->
Сечение(жилы)"""
        try:
            stage4(mark_data_ddl)  # ЖилСечение
        except:
            return \
"""Ошибка формата. stage4. Столбец(лист):
Жил(кабели) + Сечение(кабели) ->
ЖилСечение(кабели)"""
        try:
            stage5(mark_data_ddl)  # Длина
        except:
            return \
"""Ошибка формата. stage5. Столбец(лист):
Длина(кабели) -> Длина(кабели)"""
        try:
            stage6(mark_data_ddl)  # Кол.
        except:
            return \
"""Ошибка формата. stage6. Столбец(лист):
Кол.(кабели) -> все(кабели)"""
        try:
            # Кабель + ЖИЛА + Начало + Сечение +
            # Кабель + ЖИЛА + Конец  + Сечение
            stage7(mark_data_ddl)
        except:
            return \
"""Ошибка формата. stage7. Столбец(лист):
Кабель(жилы) + ЖИЛА(жилы) + Начало(жилы) + Сечение(жилы)"""
    if 'клеммы' in mark_data_ddl:
        try:
            stage8(mark_data_ddl)
        except:
            return \
"""Ошибка формата. stage8. Столбец(лист):
Текст(клеммы) + Вид(клеммы) ->
Текст(элементЗПО) + Текст(элементНПО)"""
    if 'провода' in mark_data_ddl:
        try:
            stage9(mark_data_ddl)
        except:
            return \
"""Ошибка формата. stage9. Столбец(лист):
Все(провод) ->
Адрес(провод) + Сечение(провод)"""
    try:
        # Преобразование в соответствии с форматом трансферного файла
        transfer_DDL = convert_to_transfer(mark_data_ddl)
    except:
        return \
"""Ошибка преобразования к формату трансферного файла"""
    try:
        # Обратное преобразование для экспорта
        transfer_DLL = to_dict_list_list(transfer_DDL)
    except:
        return \
"""Ошибка преобразования для сохранения"""
    try:
        # ЭКСПОРТ В ТРАНСФЕРНЫЙ ФАЙЛ
        pyexcel.save_book_as(bookdict=transfer_DLL, dest_file_name=dst)
    except:
        return \
"""Ошибка сохранения трансферного файла"""
    return 0


def gui_main():
    """Главное окно программы"""

    layout_range = 20
    layout_distance = 20

    sg.theme(GUI_THEME)  # Применение темы интерфейса

    layout_main_window = [
        [sg.Text('Производитель принтера:',
                 size=(20, 1),
                 pad=((0, 0), (10, layout_range))),
         sg.DropDown(print_programs,
                     size=(max([len(i) for i in print_programs]) + 3, 1),
                     key='#ProgramSelection',
                     enable_events=True,
                     pad=((0, layout_distance), (10, layout_range))),
         sg.Button('Установить пакет поддержки',
                   size=(26, 1),
                   key='#InstallPack',
                   pad=((0, layout_distance + 151), (10, layout_range))),
         sg.Button('Шрифт InconsolataCyr.ttf',
                   size=(24, 1),
                   key='#InstallFont',
                   pad=((0, 0), (10, layout_range)))
        ],
        [sg.Text('Файл данных:',
                 size=(20, 1),
                 pad=((0, 0), (0, layout_range))),
         sg.InputText('',
                      size=(93, 1),
                      key='#FilePath',
                      pad=((0, layout_distance + 2), (0, layout_range))),
         sg.FileBrowse('Открыть',
                       size=(8, 1),
                       key='#OpenFile',
                       target='#FilePath',
                       pad=((2, 0), (0, layout_range)))
        ],
        [sg.Button('Обработать',
                   size=(15, 2),
                   key='#ProcFile',
                   pad=((0, 33), (0, 10))),
         sg.Button('Импорт',
                   size=(9, 2),
                   key='#Import',
                   pad=((0, 473), (0, 10))),
         sg.Button('Справка',
                   size=(9, 2),
                   key='#Man',
                   pad=((0, 33), (0, 10))),
         sg.Button('Выход',
                   size=(10, 2),
                   key='#Exit',
                   pad=((0, 0), (0, 10)))
        ]
    ]

    window = sg.Window('{0}, версия {1}'.format(__progName__,
                                               __version__),
                       layout_main_window)

    while True:
        event, values = window.read()
        print(event, values)

        if event in (None, '#Exit'):
            break

        if event == '#ProgramSelection':
            if not prog_installed_check(values['#ProgramSelection']):
                prog_installed_flag = False
                sg.PopupError(
"""Программное обеспечение для печати не установлено""",
                              title='Ошибка!',
                              keep_on_top=True)
            else:
                prog_installed_flag = True
                import_file_path = os.path.join(
                      gpd('Программы',
                          'KEY',
                          values['#ProgramSelection']
                          )['data_path'],
                      gpd('Программы',
                          'KEY',
                          values['#ProgramSelection']
                          )['transfer_file_name'])
                if not pack_installed_check(values['#ProgramSelection']):
                    pack_installed_flag = False
                    sg.PopupError(
"""Пакет поддержки программы печати не установлен или повреждён.
Для установки пакета нажмите кнопку \"Установить пакет поддержки\"""",
                                  title='Ошибка!',
                                  keep_on_top=True)
                else:
                    pack_installed_flag = True

        elif event == '#InstallPack':
            if values['#ProgramSelection'] == '':
                sg.Popup(
"""Выберите производителя принтера""",
                         title='Ошибка!',
                         keep_on_top=True)
            elif not prog_installed_flag:
                sg.PopupError(
"""Программное обеспечение для печати не установлено""",
                              title='Ошибка!',
                              keep_on_top=True)
            else:
                if install_pack(values['#ProgramSelection']):
                    sg.Popup(
"""Пакет поддержки программы печати успешно установлен.""",
                             title='Успех!',
                             keep_on_top=True)
                    pack_installed_flag = True
                else:
                    sg.PopupError(
"""Не удалось установить пакет поддержки программы печати""",
                                  title='Ошибка!',
                                  keep_on_top=True)

        elif event == '#InstallFont':
            fontfile = gpd('Пути', 'KEY', 'VALUE')['font']
            if os.path.exists(fontfile):
                os.startfile(fontfile)
            else:
                sg.PopupError(
"""Не найден файл шрифта. Переустановите программу MarkV""",
                              title='Ошибка!',
                              keep_on_top=True)

        elif event == '#ProcFile':
            if values['#ProgramSelection'] == '':
                sg.Popup(
"""Выберите производителя принтера""",
                         title='Ошибка!',
                         keep_on_top=True)
            elif not prog_installed_flag:
                sg.PopupError(
"""Программное обеспечение для печати не установлено""",
                              title='Ошибка!',
                              keep_on_top=True)
            elif values['#FilePath'] == '':
                sg.Popup(
"""Выберите файл данных маркировки""",
                         title='Ошибка!',
                         keep_on_top=True)
            elif not os.path.exists(values['#FilePath']):
                sg.PopupError(
"""Указанный файл данных не найден""",
                              title='Ошибка!',
                              keep_on_top=True)
            else:
                # try:
                ret = proc_mark_file(values['#FilePath'],
                                     import_file_path)
                # except:
                    # ret = 'Неизвестная ошибка формата данных'
                if ret == 0:
                    if sg.PopupYesNo(
"""Файл данных успешно обработан.\n
Сформирован и сохранён трансферный файл.\n\n
Открыть трансферный файл для просмотра?\n""",
                          title='Успех!',
                          keep_on_top=True) == 'Yes':
                        # Открытие трансферного файла для проверки
                        os.startfile(import_file_path)
                else:
                    sg.Popup(ret, title='Ошибка!', keep_on_top=True)

        elif event == '#Import':
            if values['#ProgramSelection'] == '':
                sg.Popup(
"""Выберите производителя принтера""",
                         title='Ошибка!',
                         keep_on_top=True)
            elif not prog_installed_flag:
                sg.PopupError(
"""Программное обеспечение для печати не установлено""",
                              title='Ошибка!',
                              keep_on_top=True)
            elif not pack_installed_flag:
                sg.PopupError(
"""Пакет поддержки программы печати не установлен или повреждён.
Для установки нажмите кнопку \"Установить пакет поддержки\"""",
                              title='Ошибка!',
                              keep_on_top=True)
            elif not os.path.exists(import_file_path):
                sg.PopupError(
"""Не найден трансферный файл.
Чтобы его сформировать выберите файл данных и
нажмите кнопку \'Обработать\"""",
                              title='Ошибка!',
                              keep_on_top=True)
            else:
                prog_string = '"{0}"'.format(gpd('Программы',
                                                 'KEY',
                                                 values['#ProgramSelection']
                                                 )['program_path'])
                param_string = '"{0}"'.format(gpd('Программы',
                                                  'KEY',
                                                  values['#ProgramSelection']
                                                  )['param_path'])
                call_string = ' '.join([prog_string, param_string])
                print(call_string)
                subprocess.Popen(call_string,
                                 shell=True,
                                 stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE)

        elif event == '#Man':
            manfile = gpd('Пути', 'KEY', 'VALUE')['man']
            if os.path.exists(manfile):
                os.startfile(manfile)
            else:
                sg.PopupError(
"""Не найден файл справки. Переустановите программу MarkV""",
                              title='Ошибка!',
                              keep_on_top=True)
            # subprocess.call('Acrobat.exe /A page=3 {0}'.format(manfile))

    window.close()



if __name__ == '__main__':
    read_program_data()
    gpd = get_programm_data
    print_programs = list(gpd('Программы').keys())[1:]

    gui_main()
