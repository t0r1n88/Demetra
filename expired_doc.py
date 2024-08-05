"""
Скрипт для проверки истекающих лицензий
"""
from support_functions import *
import pandas as pd
import time
import re
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ZeroDivisionError:
        pass
    # except ValueError:
    #     return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    # except TypeError:
    #     return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'


def check_expired_docs(data_file: str, result_folder: str):
    """
    Функция для проверки истекающих документов
    :param data_file: путь к общему файлу содержащему данные студентов
    :param result_folder: итоговая папка
    :return:файл Excel
    """
    dct_df = dict() # словарь для хранения датафреймов с колонками истекающих дат
    current_date = pd.to_datetime('today', dayfirst=True)  # получаем текущую дату

    df = pd.read_excel(data_file,dtype=str)
    df.dropna(how='all',inplace=True) # очищаем от пустых строк

    lst_date_columns = [] # список для колонок с датами
    for column in df.columns:
        if 'дата' in column.lower():
            lst_date_columns.append(column)
    df[lst_date_columns] = df[lst_date_columns].apply(pd.to_datetime, errors='coerce',dayfirst=True,format='mixed')  # Приводим к типу
    df[lst_date_columns] = df[lst_date_columns].applymap(lambda x: x.strftime('%d.%m.%Y') if isinstance(x,pd.Timestamp) else x)

    # получаем список колонок, где есть сочетание Дата_окончания
    date_end_columns = [column for column in df.columns if 'Дата_окончания' in column]

    # Создаем регулярное выражение
    pattern_symbols = re.compile(r"[/*'\[\]/\\]")
    df[date_end_columns] = df[date_end_columns].apply(pd.to_datetime,errors='coerce',dayfirst=True,format='mixed') # Приводим к типу
    for idx,name_column in enumerate(date_end_columns):
        short_name_sheet = name_column.split('Дата_окончания_')[-1][:30] # Делаем короткое имя
        # очищаем названия колонок от символов */\ []''
        short_name_sheet = re.sub(pattern_symbols,'',short_name_sheet)
        temp_df = df[df[name_column].notnull()] # очищаем от пустых
        # Добавляем колонку с числом дней между текущим и окончанием срока действия документа
        temp_df['Осталось дней'] = temp_df[name_column].apply(
            lambda x: (pd.to_datetime(x,dayfirst=True,format='mixed') - current_date).days)

        temp_df[date_end_columns] = temp_df[date_end_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, pd.Timestamp) else x)
        # Фильтруем только тех у кого меньше месяца
        temp_df = temp_df[temp_df['Осталось дней'] <= 31]
        dct_df[short_name_sheet] = temp_df

    itog_wb = write_df_to_excel_expired_docs(dct_df,False)

    itog_wb = del_sheet(itog_wb,['Sheet'])

    # генерируем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    itog_wb.save(f'{result_folder}/Истекающие документы {current_time}.xlsx')












if __name__ == '__main__':
    main_file = 'data/Данные/Общий файл.xlsx'
    main_result_folder = 'data/Результат'

    check_expired_docs(main_file,main_result_folder)

    print('Lindy Booth')
