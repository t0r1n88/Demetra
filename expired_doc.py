"""
Скрипт для проверки истекающих лицензий
"""
from support_functions import *
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill
import time
from collections import Counter
import re
import os
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None


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
    # получаем список колонок, где есть сочетание Дата_окончания

    date_end_columns = [column for column in df.columns if 'Дата_окончания' in column]


    # Создаем регулярное выражение
    pattern_symbols = re.compile(r"[/*'\[\]/\\]")

    print(date_end_columns)
    df[date_end_columns] = df[date_end_columns].apply(pd.to_datetime,errors='coerce',dayfirst=True) # Приводим к типу
    for idx,name_column in enumerate(date_end_columns):
        short_name_sheet = name_column.split('Дата_окончания_')[-1][:30] # Делаем короткое имя
        # очищаем названия колонок от символов */\ []''
        short_name_sheet = re.sub(pattern_symbols,'',short_name_sheet)
        temp_df = df[df[name_column].notnull()] # очищаем от пустых
        # Добавляем колонку с числом дней между текущим и окончанием срока действия документа
        temp_df['Осталось дней'] = temp_df[name_column].apply(
            lambda x: (pd.to_datetime(x,dayfirst=True) - current_date).days)
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
