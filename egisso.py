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


def add_cols_pers_data(df:pd.DataFrame,ben_cols:list,req_cols_lst:list,name_col_ben:str):
    """
    Функция для добавления колонок с персональными данными
    :param df: полный датафрейм с данными
    :param ben_cols: список колонок относящихся к льготе
    :param req_cols_lst: наименования колонок которые нужно добавить
    :param name_col_ben: наименование льготы
    :return: датафрейм с добавленными колонками
    """
    df = df[df[name_col_ben].notna()] # убираем незаполненные строки
    ben_df = df[ben_cols] # начинаем собирать датафрейм льгот
    ben_df.insert(0,'Льгота',name_col_ben) # добавляем колонку определяющую что за льгота

    for name_column in req_cols_lst:
        if name_column in df.columns:
            ben_df[name_column] = df[name_column]
        else:
            ben_df[name_column] = f'Не найдена колонка с названием {name_column}'
    ben_df.insert(10,'Тип документа','')
    ben_df.to_excel('data/dfsf.xlsx',index=False)






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
    req_lst_personal_data_cols = ['СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан']

    # Собираем датафреймы
    for name_benefit,ben_cols in benefits_cols_dct.items():
        temp_df_full = add_cols_pers_data(df.copy(),ben_cols,req_lst_personal_data_cols,name_benefit)






if __name__ == '__main__':
    main_df = pd.read_excel('data/Свод.xlsx')

    create_part_egisso_data(main_df)


    print('Lindy Booth')
