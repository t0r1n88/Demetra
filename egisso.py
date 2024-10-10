"""
Скрипт для создания файла в котором будут содержаться частичные данные для загрузки в егиссо
Паспортные данные ,снилс фио
"""
from demetra_support_functions import *
import pandas as pd
import time
import re
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

def find_cols_benefits(lst_cols:list):
    """
    Функция для поиска трех колонок отвечающих за определенные льготы
    :param lst_cols: список колонок датафрейма
    :return: словарь вида {Название колонки со статусом льготы: [Статус, Реквизиты, Дата истечения]}
    """
    # Проверяемые строки
    status_str = 'Статус_'
    requsit_str = 'Реквизит_'
    date_str = 'Дата_'

    # Словарь для хранения найденных колонок
    ben_dct = {}

    for i in range(len(lst_cols) - 2):  # Проходим по списку, оставляя место для двух следующих элементов
        if status_str in lst_cols[i] and  requsit_str in lst_cols[i + 1] and date_str in lst_cols[i + 2]:
            ben_dct[lst_cols[i]] = [lst_cols[i],lst_cols[i + 1],lst_cols[i + 2]]

    return ben_dct





def create_part_egisso_data(df:pd.DataFrame):
    """
    Функция для создания файла xlsx  содержащего в себе данные для егиссо
    ФИО,паспортные данные, снилс, колонки со льготами
    :param df: датафрейм с данными соц паспортов
    :return: 2 файла xlsx. С данными проверки корректности заполнения и с данными
    """
    lst_cols_df = list(df.columns) # создаем список

    # ищем колонки со льготами
    benefits_cols_dct = find_cols_benefits(lst_cols_df)

    # список требуемых колонок для персональных данных
    req_lst_personal_data_cols = ['СНИЛС','','','','','','','','','','','','','','','',]





if __name__ == '__main__':
    main_df = pd.read_excel('data/Свод.xlsx')

    create_part_egisso_data(main_df)


    print('Lindy Booth')
