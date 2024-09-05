"""
Вспомогательные функции
"""
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill

from pytrovich.detector import PetrovichGenderDetector
from pytrovich.enums import NamePart, Gender, Case
from pytrovich.maker import PetrovichDeclinationMaker

def capitalize_double_name(word):
    """
    Функция для того чтобы в двойных именах и фамилиях вторая часть была также с большой буквы
    """
    lst_word = word.split('-')  # сплитим по дефису
    if len(lst_word) == 1:  # если длина списка равна 1 то это не двойная фамилия и просто возвращаем слово

        return word
    elif len(lst_word) == 2:
        first_word = lst_word[0].capitalize()  # делаем первую букву слова заглавной а остальные строчными
        second_word = lst_word[1].capitalize()
        return f'{first_word}-{second_word}'
    else:
        return 'Не удалось просклонять'


def case_lastname(maker, lastname, gender, case: Case):
    """
    Функция для обработки и склонения фамилии. Это нужно для обработки случаев двойной фамилии
    """

    lst_lastname = lastname.split('-')  # сплитим по дефису

    if len(lst_lastname) == 1:  # если длина списка равна 1 то это не двойная фамилия и просто обрабатываем слово
        case_result_lastname = maker.make(NamePart.LASTNAME, gender, case, lastname)
        return case_result_lastname
    elif len(lst_lastname) == 2:
        first_lastname = lst_lastname[0].capitalize()  # делаем первую букву слова заглавной а остальные строчными
        second_lastname = lst_lastname[1].capitalize()
        # Склоняем по отдельности
        first_lastname = maker.make(NamePart.LASTNAME, gender, case, first_lastname)
        second_lastname = maker.make(NamePart.LASTNAME, gender, case, second_lastname)

        return f'{first_lastname}-{second_lastname}'


def detect_gender(detector, lastname, firstname, middlename):
    """
    Функция для определения гендера слова
    """
    #     detector = PetrovichGenderDetector() # создаем объект детектора
    try:
        gender_result = detector.detect(lastname=lastname, firstname=firstname, middlename=middlename)
        return gender_result
    except StopIteration:  # если не удалось определить то считаем что гендер андрогинный
        return Gender.ANDROGYNOUS


def decl_on_case(fio: str, case: Case) -> str:
    """
    Функция для склонения ФИО по падежам
    """
    fio = fio.strip()  # очищаем строку от пробельных символов с начала и конца
    part_fio = fio.split()  # разбиваем по пробелам создавая список где [0] это Фамилия,[1]-Имя,[2]-Отчество

    if len(part_fio) == 3:  # проверяем на длину и обрабатываем только те что имеют длину 3 во всех остальных случаях просим просклонять самостоятельно
        maker = PetrovichDeclinationMaker()  # создаем объект класса
        lastname = part_fio[0].capitalize()  # Фамилия
        firstname = part_fio[1].capitalize()  # Имя
        middlename = part_fio[2].capitalize()  # Отчество

        # Определяем гендер для корректного склонения
        detector = PetrovichGenderDetector()  # создаем объект детектора
        gender = detect_gender(detector, lastname, firstname, middlename)
        # Склоняем

        case_result_lastname = case_lastname(maker, lastname, gender, case)  # обрабатываем фамилию
        case_result_firstname = maker.make(NamePart.FIRSTNAME, gender, case, firstname)
        case_result_firstname = capitalize_double_name(case_result_firstname)  # обрабатываем случаи двойного имени
        case_result_middlename = maker.make(NamePart.MIDDLENAME, gender, case, middlename)
        # Возвращаем результат
        result_fio = f'{case_result_lastname} {case_result_firstname} {case_result_middlename}'
        return result_fio

    else:
        return 'Проверьте количество слов, должно быть 3 разделенных пробелами слова'


def create_initials(cell, checkbox, space):
    """
    Функция для создания инициалов
    """
    lst_fio = cell.split(' ')  # сплитим по пробелу
    if len(lst_fio) == 3:  # проверяем на стандартный размер в 3 слова иначе ничего не меняем
        if checkbox == 'ФИ':
            if space == 'без пробела':
                # возвращаем строку вида Иванов И.И.
                return f'{lst_fio[0]} {lst_fio[1][0].upper()}.{lst_fio[2][0].upper()}.'
            else:
                # возвращаем строку с пробелом после имени Иванов И. И.
                return f'{lst_fio[0]} {lst_fio[1][0].upper()}. {lst_fio[2][0].upper()}.'

        else:
            if space == 'без пробела':
                # И.И. Иванов
                return f'{lst_fio[1][0].upper()}.{lst_fio[2][0].upper()}. {lst_fio[0]}'
            else:
                # И. И. Иванов
                return f'{lst_fio[1][0].upper()}. {lst_fio[2][0].upper()}. {lst_fio[0]}'
    else:
        return cell

def declension_fio_by_case(df:pd.DataFrame):
    """
    Функция для склонения ФИО по падежам и создания инициалов
    :param df: датафрейм в который надо добавить инициалы и падежи

    :return: файл Excel в котором после колонки fio_column добавляется 29 колонок с падежами
    """
    fio_column = 'ФИО'
    df[fio_column] = df[fio_column].astype(str) # делаем на всякий случай строковой

    temp_df = pd.DataFrame()  # временный датафрейм для хранения колонок просклоненных по падежам

    # Получаем номер колонки с фио которые нужно обработать
    lst_columns = list(df.columns)  # Превращаем в список
    index_fio_column = lst_columns.index(fio_column)  # получаем индекс

    # Обрабатываем nan значения и те которые обозначены пробелом
    df[fio_column].fillna('Не заполнено', inplace=True)
    df[fio_column] = df[fio_column].apply(lambda x: x.strip())
    df[fio_column] = df[fio_column].apply(
        lambda x: x if x else 'Не заполнено')  # Если пустая строка то заменяем на значение Не заполнено

    temp_df['Родительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.GENITIVE))
    temp_df['Дательный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.DATIVE))
    temp_df['Винительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.ACCUSATIVE))
    temp_df['Творительный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.INSTRUMENTAL))
    temp_df['Предложный_падеж'] = df[fio_column].apply(lambda x: decl_on_case(x, Case.PREPOSITIONAL))
    temp_df['Фамилия_инициалы'] = df[fio_column].apply(lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Инициалы_фамилия'] = df[fio_column].apply(lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Фамилия_инициалы_пробел'] = df[fio_column].apply(lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_пробел'] = df[fio_column].apply(lambda x: create_initials(x, 'ИФ', 'пробел'))

    # Создаем колонки для склонения фамилий с иницалами родительный падеж
    temp_df['Фамилия_инициалы_род_падеж'] = temp_df['Родительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Фамилия_инициалы_род_падеж_пробел'] = temp_df['Родительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_род_падеж'] = temp_df['Родительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Инициалы_фамилия_род_падеж_пробел'] = temp_df['Родительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'пробел'))

    # Создаем колонки для склонения фамилий с иницалами дательный падеж
    temp_df['Фамилия_инициалы_дат_падеж'] = temp_df['Дательный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Фамилия_инициалы_дат_падеж_пробел'] = temp_df['Дательный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_дат_падеж'] = temp_df['Дательный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Инициалы_фамилия_дат_падеж_пробел'] = temp_df['Дательный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'пробел'))

    # Создаем колонки для склонения фамилий с иницалами винительный падеж
    temp_df['Фамилия_инициалы_вин_падеж'] = temp_df['Винительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Фамилия_инициалы_вин_падеж_пробел'] = temp_df['Винительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_вин_падеж'] = temp_df['Винительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Инициалы_фамилия_вин_падеж_пробел'] = temp_df['Винительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'пробел'))

    # Создаем колонки для склонения фамилий с иницалами творительный падеж
    temp_df['Фамилия_инициалы_твор_падеж'] = temp_df['Творительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Фамилия_инициалы_твор_падеж_пробел'] = temp_df['Творительный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_твор_падеж'] = temp_df['Творительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Инициалы_фамилия_твор_падеж_пробел'] = temp_df['Творительный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'пробел'))
    # Создаем колонки для склонения фамилий с иницалами предложный падеж
    temp_df['Фамилия_инициалы_пред_падеж'] = temp_df['Предложный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'без пробела'))
    temp_df['Фамилия_инициалы_пред_падеж_пробел'] = temp_df['Предложный_падеж'].apply(
        lambda x: create_initials(x, 'ФИ', 'пробел'))
    temp_df['Инициалы_фамилия_пред_падеж'] = temp_df['Предложный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'без пробела'))
    temp_df['Инициалы_фамилия_пред_падеж_пробел'] = temp_df['Предложный_падеж'].apply(
        lambda x: create_initials(x, 'ИФ', 'пробел'))

    # Вставляем получившиеся колонки в конец датафрейма
    # Падежи
    df['Родительный_падеж'] = temp_df['Родительный_падеж']
    df['Дательный_падеж'] = temp_df['Дательный_падеж']
    df['Винительный_падеж'] = temp_df['Винительный_падеж']
    df['Творительный_падеж'] = temp_df['Творительный_падеж']
    df['Предложный_падеж'] = temp_df['Предложный_падеж']
    # Инициалы
    df['Фамилия_инициалы'] = temp_df['Фамилия_инициалы']
    df['Инициалы_фамилия'] = temp_df['Инициалы_фамилия']
    df['Фамилия_инициалы_пробел'] = temp_df['Фамилия_инициалы_пробел']
    df['Инициалы_фамилия_пробел'] = temp_df['Инициалы_фамилия_пробел']

    # Добавляем колонки с склонениями инициалов родительный падеж
    df['Фамилия_инициалы_род_падеж'] = temp_df['Фамилия_инициалы_род_падеж']
    df['Фамилия_инициалы_род_падеж_пробел'] = temp_df['Фамилия_инициалы_род_падеж_пробел']
    df['Инициалы_фамилия_род_падеж'] = temp_df['Инициалы_фамилия_род_падеж']
    df['Инициалы_фамилия_род_падеж_пробел'] = temp_df['Инициалы_фамилия_род_падеж_пробел']

    # Добавляем колонки с склонениями инициалов дательный падеж
    df['Фамилия_инициалы_дат_падеж'] = temp_df['Фамилия_инициалы_дат_падеж']
    df['Фамилия_инициалы_дат_падеж_пробел'] = temp_df['Фамилия_инициалы_дат_падеж_пробел']
    df['Инициалы_фамилия_дат_падеж'] = temp_df['Инициалы_фамилия_дат_падеж']
    df['Инициалы_фамилия_дат_падеж_пробел'] = temp_df['Инициалы_фамилия_дат_падеж_пробел']

    # Добавляем колонки с склонениями инициалов винительный падеж
    df['Фамилия_инициалы_вин_падеж'] = temp_df['Фамилия_инициалы_вин_падеж']
    df['Фамилия_инициалы_вин_падеж_пробел'] = temp_df['Фамилия_инициалы_вин_падеж_пробел']
    df['Инициалы_фамилия_вин_падеж'] = temp_df['Инициалы_фамилия_вин_падеж']
    df['Инициалы_фамилия_вин_падеж_пробел'] = temp_df['Инициалы_фамилия_вин_падеж_пробел']

    # Добавляем колонки с склонениями инициалов творительный падеж
    df['Фамилия_инициалы_твор_падеж'] = temp_df['Фамилия_инициалы_твор_падеж']
    df['Фамилия_инициалы_твор_падеж_пробел'] = temp_df['Фамилия_инициалы_твор_падеж_пробел']
    df['Инициалы_фамилия_твор_падеж'] = temp_df['Инициалы_фамилия_твор_падеж']
    df['Инициалы_фамилия_твор_падеж_пробел'] = temp_df['Инициалы_фамилия_твор_падеж_пробел']

    # Добавляем колонки с склонениями инициалов предложный падеж
    df['Фамилия_инициалы_пред_падеж'] = temp_df['Фамилия_инициалы_пред_падеж']
    df['Фамилия_инициалы_пред_падеж_пробел'] = temp_df['Фамилия_инициалы_пред_падеж_пробел']
    df['Инициалы_фамилия_пред_падеж'] = temp_df['Инициалы_фамилия_пред_падеж']
    df['Инициалы_фамилия_пред_падеж_пробел'] = temp_df['Инициалы_фамилия_пред_падеж_пробел']

    return df

def write_df_to_excel(dct_df: dict, write_index: bool) -> openpyxl.Workbook:
    """
    Функция для записи датафрейма в файл Excel
    :param dct_df: словарь где ключе это название создаваемого листа а значение датафрейм который нужно записать
    :param write_index: нужно ли записывать индекс датафрейма True or False
    :return: объект Workbook с записанными датафреймами
    """
    wb = openpyxl.Workbook()  # создаем файл
    count_index = 0  # счетчик индексов создаваемых листов
    for name_sheet, df in dct_df.items():
        wb.create_sheet(title=name_sheet, index=count_index)  # создаем лист
        # записываем данные в лист
        none_check = None  # чекбокс для проверки наличия пустой первой строки, такое почему то иногда бывает
        for row in dataframe_to_rows(df, index=write_index, header=True):
            if len(row) == 1 and not row[0]:  # убираем пустую строку
                none_check = True
                wb[name_sheet].append(row)
            else:
                wb[name_sheet].append(row)
        if none_check:
            wb[name_sheet].delete_rows(2)

        # ширина по содержимому
        # сохраняем по ширине колонок
        for column in wb[name_sheet].columns:
            max_length = 0
            column_name = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            wb[name_sheet].column_dimensions[column_name].width = adjusted_width
        count_index += 1

    return wb


def write_df_to_excel_report_brit(dct_df: dict, write_index: bool) -> openpyxl.Workbook:
    """
    Функция для записи датафрейма в файл Excel отчета по стандарту БРИТ
    :param dct_df: словарь где ключе это название создаваемого листа а значение датафрейм который нужно записать
    :param write_index: нужно ли записывать индекс датафрейма True or False
    :return: объект Workbook с записанными датафреймами
    """
    wb = openpyxl.Workbook()  # создаем файл
    count_index = 0  # счетчик индексов создаваемых листов
    for name_sheet, df in dct_df.items():
        wb.create_sheet(title=name_sheet, index=count_index)  # создаем лист
        # записываем данные в лист
        none_check = None  # чекбокс для проверки наличия пустой первой строки, такое почему то иногда бывает
        for row in dataframe_to_rows(df, index=write_index, header=True):
            if len(row) == 1 and not row[0]:  # убираем пустую строку
                none_check = True
                wb[name_sheet].append(row)
            else:
                wb[name_sheet].append(row)
        if none_check:
            wb[name_sheet].delete_rows(2)

        # ширина по содержимому
        # сохраняем по ширине колонок
        for column in wb[name_sheet].columns:
            max_length = 0
            column_name = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            wb[name_sheet].column_dimensions[column_name].width = adjusted_width
        count_index += 1
        column_number = 0  # номер колонки в которой ищем слово Статус_
        # Создаем  стиль шрифта и заливки
        font = Font(color='FF000000')  # Черный цвет
        fill = PatternFill(fill_type='solid', fgColor='ffa500')  # Оранжевый цвет
        for row in wb[name_sheet].iter_rows(min_row=1, max_row=wb[name_sheet].max_row,
                                            min_col=column_number, max_col=df.shape[1] + 1):  # Перебираем строки
            if 'Итого' in str(row[column_number].value):  # делаем ячейку строковой и проверяем наличие слова Статус_
                for cell in row:  # применяем стиль если условие сработало
                    cell.font = font
                    cell.fill = fill

    return wb


def write_df_to_excel_expired_docs(dct_df: dict, write_index: bool) -> openpyxl.Workbook:
    """
    Функция для записи датафрейма в файл Excel
    :param dct_df: словарь где ключе это название создаваемого листа а значение датафрейм который нужно записать
    :param write_index: нужно ли записывать индекс датафрейма True or False
    :return: объект Workbook с записанными датафреймами
    """
    wb = openpyxl.Workbook()  # создаем файл
    count_index = 0  # счетчик индексов создаваемых листов
    for name_sheet, df in dct_df.items():
        wb.create_sheet(title=name_sheet, index=count_index)  # создаем лист
        # записываем данные в лист
        none_check = None  # чекбокс для проверки наличия пустой первой строки, такое почему то иногда бывает
        for row in dataframe_to_rows(df, index=write_index, header=True):
            if len(row) == 1 and not row[0]:  # убираем пустую строку
                none_check = True
                wb[name_sheet].append(row)
            else:
                wb[name_sheet].append(row)
        if none_check:
            wb[name_sheet].delete_rows(2)

        # ширина по содержимому
        # сохраняем по ширине колонок
        for column in wb[name_sheet].columns:
            max_length = 0
            column_name = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            wb[name_sheet].column_dimensions[column_name].width = adjusted_width
        # Красим в зависимости от цвета
        last_column_number = wb.active.max_column  # номер колонки в которой ищем слово Осталось дней
        # Создаем  стиль шрифта и заливки
        font = Font(color='FF000000')  # Черный цвет
        # Месяц
        fill_week = PatternFill(fill_type='solid', fgColor='ff0000')  # Красный цвет
        fill_two_week = PatternFill(fill_type='solid', fgColor='ffa500')  # Оранжевый цвет
        fill_month = PatternFill(fill_type='solid', fgColor='ffff00')  # Желтый цвет
        for row in wb[name_sheet].iter_rows(min_row=2, max_row=wb[name_sheet].max_row,
                                            min_col=0,
                                            max_col=last_column_number):  # Перебираем строки
            if int(row[last_column_number-1].value) <= 7:
                for cell in row:
                    cell.font = font
                    cell.fill = fill_week
            elif int(row[last_column_number-1].value) <= 14:
                for cell in row:
                    cell.font = font
                    cell.fill = fill_two_week
            else:
                for cell in row:
                    cell.font = font
                    cell.fill = fill_month


        count_index += 1

    return wb


def del_sheet(wb: openpyxl.Workbook, lst_name_sheet: list) -> openpyxl.Workbook:
    """
    Функция для удаления лишних листов из файла
    :param wb: объект таблицы
    :param lst_name_sheet: список удаляемых листов
    :return: объект таблицы без удаленных листов
    """
    for del_sheet in lst_name_sheet:
        if del_sheet in wb.sheetnames:
            del wb[del_sheet]

    return wb