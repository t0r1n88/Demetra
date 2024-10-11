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
    df[name_col_ben] = df[name_col_ben].replace('нет',None) # подчищаем заполненные нет
    df = df[df[name_col_ben].notna()] # убираем незаполненные строки
    ben_df = df[ben_cols] # начинаем собирать датафрейм льгот
    ben_df.columns = ['Статус льготы','Реквизиты','Дата окончания льготы',]
    name_col_ben = name_col_ben.replace('Статус_','')
    ben_df.insert(0,'Льгота',name_col_ben) # добавляем колонку определяющую, что за льгота

    for name_column in req_cols_lst:
        if name_column in df.columns:
            ben_df[name_column] = df[name_column]
        else:
            ben_df[name_column] = f'В исходном файле не найдена колонка с названием {name_column}'
    ben_df.insert(10,'Тип документа','')

    return ben_df

def check_simple_str_column(value,error_str:str):
    """
    Функция для проверки на заполнение ячейки для простой колонки с текстом не требующим дополнительной проверки
    :param value: значение ячейки
    :param error_str: сообщение об ошибки
    """
    if pd.isna(value):
        return error_str
    else:
        return value

def processing_snils(value):
    result = re.findall(r'\d',value)
    if len(result) == 11:
        # проверяем на лидирующий ноль
        out_str = ''.join(result)
        if out_str.startswith('0'):
            return out_str
        else:
            return int(out_str)
    else:
        return f'Ошибка: В СНИЛС должно быть 11 цифр а в ячейке {len(result)} цифр(ы) - {value}'


def check_error_ben(df:pd.DataFrame):
    """
    Функция для проверки правильности данных
    :param df:датафрейм с данными по одной льготе
    :return:2 датафрейма  один без ошибок и второй где указаны ошибки
    """
    # Базовые датафреймы
    clean_df = pd.DataFrame(columns=['Льгота','Статус льготы','Реквизиты','Дата окончания льготы','Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан'])
    error_df = pd.DataFrame(columns=['Льгота','Статус льготы','Реквизиты','Дата окончания льготы','Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан'])

    checked_simple_cols = ['СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан']

    df[checked_simple_cols] = df[checked_simple_cols].applymap(lambda x:check_simple_str_column(x,'Не заполнено'))
    df['СНИЛС'] = df['СНИЛС'].apply(processing_snils)

    df.to_excel('data/tres.xlsx')






def create_part_egisso_data(df:pd.DataFrame):
    """
    Функция для создания файла xlsx  содержащего в себе данные для егиссо
    ФИО,паспортные данные, снилс, колонки со льготами
    :param df: датафрейм с данными соц паспортов
    :return: 2 файла xlsx. С данными проверки корректности заполнения и с данными
    """
    main_df = pd.DataFrame(columns=['Льгота','Статус льготы','Реквизиты','Дата окончания льготы','Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан'])
    lst_cols_df = list(df.columns) # создаем список

    # ищем колонки со льготами
    benefits_cols_dct = find_cols_benefits(lst_cols_df)

    # список требуемых колонок для персональных данных
    req_lst_personal_data_cols = ['Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан']

    # Собираем датафреймы
    for name_benefit,ben_cols in benefits_cols_dct.items():
        if name_benefit == 'Статус_Уровень_здоровья':
            health_df = df.copy() # костыль из-за того что в статус уровень здоровья для здоровых тоже указаны значения
            health_df['Статус_Уровень_здоровья'] = health_df['Статус_Уровень_здоровья'].fillna('доров')
            health_df = health_df[~health_df['Статус_Уровень_здоровья'].str.contains('доров')]
            temp_df_full = add_cols_pers_data(health_df,ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе
        else:
            temp_df_full = add_cols_pers_data(df.copy(),ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе

        check_error_ben(temp_df_full)
        main_df = pd.concat([main_df,temp_df_full])







if __name__ == '__main__':
    main_df = pd.read_excel('data/Свод.xlsx')

    create_part_egisso_data(main_df)


    print('Lindy Booth')
