"""
Скрипт для создания  отчета по социальному паспорту студента БРИТ
"""
from demetra_support_functions import (write_df_to_excel,write_df_to_excel_report_brit,del_sheet,declension_fio_by_case,
                                       extract_parameters_egisso,write_df_big_dct_to_excel)
from demetra_processing_date import proccessing_date
from egisso import create_part_egisso_data, create_full_egisso_data
from tkinter import messagebox

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
import time
import datetime
from collections import Counter
import re
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action="ignore", category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import sys
import locale

class NotColumn(Exception):
    """
    Исключение для обработки случая когда отсутствуют нужные колонки
    """
    pass

class NotGoodSheet(Exception):
    """
    Исключение для случая когда ни один лист не подхожит под эталон
    """
    pass

class ExceedingQuantity(Exception):
    """
    Исключение для случаев когда числа уникальных значений больше 255
    """
    pass

def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')
def create_value_str(df:pd.DataFrame,name_column:str,target_name_column:str,dct_str:dict)->pd.DataFrame:
    """
    Функция для формирования строки нужного формата с использованием переменных
    :param df:датафрейм
    :param name_column:название колонки для значений которой нужно произвести подсчет
    :param target_name_column: название колонки по которой будет производится подсчет
    :param dct_str:словарь с параметрами
    :return:датафрейм
    """
    temp_counts = df[name_column].value_counts()  # делаем подсчет
    new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
    new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
    new_value_df['Показатель'] = new_value_df['Показатель'].astype(str)
    new_value_df.sort_values(by='Показатель', inplace=True)
    for idx,row in enumerate(new_value_df.iterrows()):
        name_op = row[1].values[0] # получаем название ОП
        temp_df = df[df[name_column] == name_op] # отфильтровываем по названию ОП
        quantity_study_student = temp_df[temp_df[target_name_column] == dct_str['Обучается']].shape[0]  # со статусом Обучается
        quantity_academ_student = temp_df[temp_df[target_name_column].str.contains(dct_str['Академ'])].shape[
            0]
        quantity_not_status_student = temp_df[temp_df[target_name_column].str.contains(dct_str['Не указан статус'])].shape[
            0]
        quantity_except_deducted = temp_df[~temp_df[target_name_column].str.contains('Отчислен')].shape[
            0]
        out_str = f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student},' \
                  f' Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)'
        new_value_df.iloc[idx,1] = out_str # присваиваем значение

    return new_value_df

def count_value(group: pd.Series, target_value: str):
        """
        Функция для группировки по конкретному значению
        Возвращает количество значений target_value в группе
        """
        # считаем сколько значений подходят по условие
        count_group = group.str.contains(target_value)
        return sum(count_group)

def convert_number(value):
    """
    Функция для конвертации в float значений колонок содержащих в названии Подсчет_
    :param value: значение
    :return: число в формате float
    """
    try:
        return float(value)
    except:
        return 0


def convert_to_date(value):
    """
    Функция для конвертации строки в текст
    :param value: значение для конвертации
    :return:
    """
    try:
        if value == 'Нет статуса':
            return None
        else:
            date_value  = datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
            return date_value
    except ValueError:
        result = re.search(r'^\d{2}\.\d{2}\.\d{4}$',value)
        if result:
            return datetime.datetime.strptime(result.group(0), '%d.%m.%Y')
        else:
            return f'Некорректный формат даты - {value}'
    except:
        return None





def create_report_brit(df:pd.DataFrame,path_end_folder:str)->None:
    """
    Функция для создания отчета по стандарту БРИТ
    :param df: копия общего датафрейма
    :param path_end_folder: куда сохранять результаты
    """
    dct_name_sheet = dict()  # словарь где ключ это названия листа а значение датафрейм на основе которого был произведен подсчет

    df.fillna('Нет статуса', inplace=True)  # заполняем Наны

    group_main_df = pd.DataFrame(index=list(df['Файл'].unique()))

    # Считаем общее количество студентов
    study_df = df[df['Статус_Учёба'] == 'Обучается']
    dct_name_sheet['Обучается'] = study_df  # добавляем в словарь
    study_df_group_df = study_df.groupby(by=['Файл']).agg({'ФИО': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(study_df_group_df)  # добавляем в свод
    group_main_df.rename(columns={'ФИО': 'Обучается'}, inplace=True)

    # Считаем количество студентов в академе
    akadem_df = df[df['Статус_Учёба'].str.contains('Академ')]
    dct_name_sheet['Академ'] = akadem_df  # добавляем в словарь
    akadem_df_group_df = akadem_df.groupby(by=['Файл']).agg({'ФИО': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(akadem_df_group_df)  # добавляем в свод
    group_main_df.rename(columns={'ФИО': 'Академ. отпуск'}, inplace=True)



    # Совершеннолетние
    maturity_df = df[df['Совершеннолетие'] == 'совершеннолетний']
    dct_name_sheet['Совершеннолетние'] = maturity_df  # добавляем в словарь
    maturity_df_group_df = maturity_df.groupby(by=['Файл']).agg({'Совершеннолетие': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(maturity_df_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Совершеннолетие': 'Совершеннолетние'}, inplace=True)
    # Несовершеннолетние
    not_maturity_df = df[df['Совершеннолетие'] == 'несовершеннолетний']
    dct_name_sheet['Несовершеннолетние'] = not_maturity_df  # добавляем в словарь
    not_maturity_df_group_df = not_maturity_df.groupby(by=['Файл']).agg({'Совершеннолетие': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(not_maturity_df_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Совершеннолетие': 'Несовершеннолетние'}, inplace=True)

    # Создаем датафрейм с сиротами
    orphans_df = df[df['Статус_Сиротство'].isin(['гособеспечение + постинтернатное сопровождение',
                                                 'дети-сироты, находящиеся на полном государственном обеспечении',
                                                 'дети-сироты, находящиеся под опекой'])]

    dct_name_sheet['Сироты'] = orphans_df  # добавляем в словарь
    orphans_group_df = orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(orphans_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Сиротство': 'Дети-сироты'}, inplace=True)


    # Создаем датафрейм с инвалидами
    invalid_df = df[df['Статус_Уровень_здоровья'].isin(['Инвалид детства', 'Инвалид 1,2,3, группы'])]

    dct_name_sheet['Инвалиды'] = invalid_df  # добавляем в словарь
    invalid_group_df = invalid_df.groupby(by=['Файл']).agg({'Статус_Уровень_здоровья': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(invalid_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Инвалиды'}, inplace=True)

    # Создаем датафрейм с совершеннолетними инвалидами
    maturity_invalid_df = invalid_df[invalid_df['Совершеннолетие'] == 'совершеннолетний']
    dct_name_sheet['Инвалиды_сов'] = maturity_invalid_df  # добавляем в словарь
    maturity_invalid_group_df = maturity_invalid_df.groupby(by=['Файл']).agg({'Статус_Уровень_здоровья': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(maturity_invalid_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Инвалиды совер-ние'}, inplace=True)

    # Создаем датафрейм с несовершеннолетними инвалидами
    not_maturity_invalid_df = invalid_df[invalid_df['Совершеннолетие'] == 'несовершеннолетний']
    dct_name_sheet['Инвалиды_несов'] = not_maturity_invalid_df  # добавляем в словарь
    not_maturity_invalid_group_df = not_maturity_invalid_df.groupby(by=['Файл']).agg({'Статус_Уровень_здоровья': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(not_maturity_invalid_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Инвалиды несов-ние'}, inplace=True)


    # Создаем датафрейм с получателями социальной стипендии
    soc_benefit_df = df[df['Статус_Соц_стипендия'].isin(['да'])]
    dct_name_sheet['Соц. стипендия Все'] = soc_benefit_df  # добавляем в словарь

    # получаем малоимущих
    poor_soc_benefit_df = soc_benefit_df[soc_benefit_df['Статус_Соц_положение_семьи'].isin(['Малоимущая'])]
    dct_name_sheet['Соц. стипендия малоим.'] = poor_soc_benefit_df
    poor_soc_benefit_group_df = poor_soc_benefit_df.groupby(by=['Файл']).agg({'Статус_Соц_положение_семьи': 'count'})
    group_main_df = group_main_df.join(poor_soc_benefit_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Соц_положение_семьи': 'Соц. стипендия малоимущие'}, inplace=True)

    # получаем детей сирот
    orphans_soc_benefit_df = soc_benefit_df[soc_benefit_df['Статус_Сиротство'].isin(
        ['гособеспечение + постинтернатное сопровождение',
         'дети-сироты, находящиеся на полном государственном обеспечении',
         'дети-сироты, находящиеся под опекой'])]
    dct_name_sheet['Соц. стипендия сироты'] = orphans_soc_benefit_df
    orphans_soc_benefit_group_df = orphans_soc_benefit_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_main_df = group_main_df.join(orphans_soc_benefit_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Сиротство': 'Соц. стипендия дети-сироты'}, inplace=True)

    # получаем инвалидов
    invalid_soc_benefit_df = soc_benefit_df[
        soc_benefit_df['Статус_Уровень_здоровья'].isin(['Инвалид детства', 'Инвалид 1,2,3, группы'])]
    dct_name_sheet['Соц. стипендия инвалиды'] = invalid_soc_benefit_df
    invalid_soc_benefit_group_df = invalid_soc_benefit_df.groupby(by=['Файл']).agg({'Статус_Уровень_здоровья': 'count'})
    group_main_df = group_main_df.join(invalid_soc_benefit_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Соц. стипендия инвалиды'}, inplace=True)

    # Создаем датафрейм с получателями бесплатного питания
    eating_df = df[df['Статус_Питание'].isin(['получает компенсацию за питание', 'питается в ПОО','получает компенсацию за питание + питается в ПОО'])]
    dct_name_sheet['Питание все'] = eating_df  # добавляем в словарь

    # получаем малоимущих
    poor_eating_df = eating_df[eating_df['Статус_Соц_положение_семьи'].isin(['Малоимущая'])]
    dct_name_sheet['Питание малоим.'] = poor_eating_df
    poor_eating_group_df = poor_eating_df.groupby(by=['Файл']).agg({'Статус_Соц_положение_семьи': 'count'})
    group_main_df = group_main_df.join(poor_eating_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Соц_положение_семьи': 'Питание малоимущие'}, inplace=True)

    # получаем сирот
    orphans_eating_df = eating_df[eating_df['Статус_Сиротство'].isin(['гособеспечение + постинтернатное сопровождение',
                                                                      'дети-сироты, находящиеся на полном государственном обеспечении',
                                                                      'дети-сироты, находящиеся под опекой'])]
    dct_name_sheet['Питание сироты'] = orphans_eating_df
    orphans_eating_group_df = orphans_eating_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_main_df = group_main_df.join(orphans_eating_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Сиротство': 'Питание сироты'}, inplace=True)

    # получаем инвалидов
    invalid_eating_df = eating_df[
        eating_df['Статус_Уровень_здоровья'].isin(['Инвалид детства', 'Инвалид 1,2,3, группы'])]
    dct_name_sheet['Питание инвалиды'] = invalid_eating_df
    invalid_eating_group_df = invalid_eating_df.groupby(by=['Файл']).agg({'Статус_Уровень_здоровья': 'count'})
    group_main_df = group_main_df.join(invalid_eating_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Питание инвалиды'}, inplace=True)

    # получаем СВО
    svo_eating_df = eating_df[eating_df['Статус_Родитель_СВО'].isin(['да'])]
    dct_name_sheet['Питание СВО'] = svo_eating_df
    svo_eating_group_df = svo_eating_df.groupby(by=['Файл']).agg({'Статус_Родитель_СВО': 'count'})
    group_main_df = group_main_df.join(svo_eating_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Родитель_СВО': 'Питание СВО'}, inplace=True)

    # Создаем датафрейм с проживающими в общежитии
    dormitory_df = df[df['Статус_Общежитие'].isin(['да'])]
    dct_name_sheet['Общежитие все'] = dormitory_df  # добавляем в словарь
    dormitory_group_df = dormitory_df.groupby(by=['Файл']).agg({'Статус_Общежитие': 'count'})
    group_main_df = group_main_df.join(dormitory_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Общежитие': 'Общежитие Всего'}, inplace=True)

    # получаем не сирот и сохраняем для списка
    not_orphans_dormitory_df = dormitory_df[~dormitory_df['Статус_Сиротство'].isin(
        ['гособеспечение + постинтернатное сопровождение',
         'дети-сироты, находящиеся на полном государственном обеспечении',
         'дети-сироты, находящиеся под опекой'])]
    dct_name_sheet['Общежитие кроме сирот'] = not_orphans_dormitory_df

    # получаем сирот
    orphans_dormitory_df = dormitory_df[dormitory_df['Статус_Сиротство'].isin(
        ['гособеспечение + постинтернатное сопровождение',
         'дети-сироты, находящиеся на полном государственном обеспечении',
         'дети-сироты, находящиеся под опекой'])]
    dct_name_sheet['Общежитие сироты'] = orphans_dormitory_df
    orphans_dormitory_group_df = orphans_dormitory_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_main_df = group_main_df.join(orphans_dormitory_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Сиротство': 'Общежитие сироты'}, inplace=True)

    # Считаем выпуск текущего год
    release_df = df[df['Статус_Выпуск'].isin(['да'])]
    dct_name_sheet['Выпуск текущий год'] = release_df  # добавляем в словарь

    # Считаем сирот
    release_orphans_df = release_df[release_df['Статус_Сиротство'].isin(
        ['гособеспечение + постинтернатное сопровождение',
         'дети-сироты, находящиеся на полном государственном обеспечении',
         'дети-сироты, находящиеся под опекой'])]

    dct_name_sheet['Выпуск сироты'] = release_orphans_df  # добавляем в словарь
    release_orphans_group_df = release_orphans_df.groupby(by=['Файл']).agg(
        {'Статус_Сиротство': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(release_orphans_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Сиротство': 'Выпуск сироты'}, inplace=True)

    # Считаем инвалидов
    release_invalid_df = release_df[
        release_df['Статус_Уровень_здоровья'].isin(['Инвалид детства', 'Инвалид 1,2,3, группы'])]

    dct_name_sheet['Выпуск инвалиды'] = release_invalid_df  # добавляем в словарь
    release_invalid_group_df = release_invalid_df.groupby(by=['Файл']).agg(
        {'Статус_Уровень_здоровья': 'count'})  # создаем базовый
    group_main_df = group_main_df.join(release_invalid_group_df)  # добавляем в свод
    group_main_df.rename(columns={'Статус_Уровень_здоровья': 'Выпуск инвалиды'}, inplace=True)

    group_main_df.fillna(0, inplace=True)  # заполняем наны
    group_main_df = group_main_df.astype(int)  # приводим к инту
    sum_row = group_main_df.sum(axis=0)  # суммируем колонки
    group_main_df.loc['Итого'] = sum_row  # добавляем суммирующую колонку


    group_orphans_main_df = pd.DataFrame(index=list(df['Файл'].unique()))  # Базовый датафрейм для сирот

    # Считаем общее количество
    all_orphans_group_df = orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})  # создаем базовый
    group_orphans_main_df = group_orphans_main_df.join(all_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Всего'}, inplace=True)

    # Считаем сирот совершеннолетних
    maturity_orphans_df = orphans_df[orphans_df['Совершеннолетие'] == 'совершеннолетний']
    dct_name_sheet['Сироты совер-ние'] = maturity_orphans_df
    maturity_orphans_group_df = maturity_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(maturity_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Сироты совершеннолетние'}, inplace=True)

    # Считаем сирот несовершеннолетних
    not_maturity_orphans_df = orphans_df[orphans_df['Совершеннолетие'] == 'несовершеннолетний']
    dct_name_sheet['Сироты несовер-ние'] = not_maturity_orphans_df
    not_maturity_orphans_group_df = not_maturity_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(not_maturity_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Сироты несовершеннолетние'}, inplace=True)

    # считаем академ
    akadem_orphans_df = orphans_df[orphans_df['Статус_Учёба'].str.contains('Академический отпуск')]
    dct_name_sheet['Сироты академ'] = akadem_orphans_df
    akadem_orphans_group_df = akadem_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(akadem_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Из них в академическом отпуске'}, inplace=True)

    # считаем постинтернатное сопровождение
    postinternat_orphans_df = orphans_df[orphans_df['Статус_Сиротство'].str.contains('постинтернатное сопровождение')]
    dct_name_sheet['Сироты постинтернат'] = postinternat_orphans_df
    postinternat_orphans_group_df = postinternat_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(postinternat_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Из них постинтернат'}, inplace=True)

    # считаем сирот на гособеспечении
    gos_orphans_df = orphans_df[orphans_df['Статус_Сиротство'].isin(['гособеспечение + постинтернатное сопровождение',
                                                                     'дети-сироты, находящиеся на полном государственном обеспечении'
                                                                     ])]
    dct_name_sheet['Сироты гособеспечение'] = gos_orphans_df
    gos_orphans_group_df = gos_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(gos_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Из них на гособеспечении'}, inplace=True)

    # считаем сирот с опекой
    custody_orphans_df = orphans_df[orphans_df['Статус_Сиротство'].isin(['дети-сироты, находящиеся под опекой'])]
    dct_name_sheet['Сироты опека'] = custody_orphans_df
    custody_orphans_group_df = custody_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(custody_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Из них под опекой'}, inplace=True)

    # считаем сирот получающих питание
    all_eating_orphans_df = orphans_df[
        orphans_df['Статус_Питание'].isin(['питается в ПОО', 'получает компенсацию за питание',
                                           'получает компенсацию за питание + питается в ПОО'])]
    dct_name_sheet['Сироты питание все'] = all_eating_orphans_df
    all_eating_orphans_group_df = all_eating_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(all_eating_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Питается всего'}, inplace=True)

    # считаем сирот получающих питание в брит
    brit_eating_orphans_df = orphans_df[orphans_df['Статус_Питание'].isin(['питается в ПОО'])]
    dct_name_sheet['Сироты питание ПОО'] = brit_eating_orphans_df
    brit_eating_orphans_group_df = brit_eating_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(brit_eating_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Питается в ПОО'}, inplace=True)

    # считаем сирот получающих компенсацию
    compens_eating_orphans_df = orphans_df[orphans_df['Статус_Питание'].isin(['получает компенсацию за питание'])]
    dct_name_sheet['Сироты питание компенсация'] = compens_eating_orphans_df
    compens_eating_orphans_group_df = compens_eating_orphans_df.groupby(by=['Файл']).agg({'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(compens_eating_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'Получает компенсацию'}, inplace=True)

    # считаем сирот получающих компенсацию + питание в брит
    brit_compens_eating_orphans_df = orphans_df[
        orphans_df['Статус_Питание'].isin(['получает компенсацию за питание + питается в ПОО'])]
    dct_name_sheet['Сироты питание ПОО+компенсация'] = brit_compens_eating_orphans_df
    brit_compens_eating_orphans_group_df = brit_compens_eating_orphans_df.groupby(by=['Файл']).agg(
        {'Статус_Сиротство': 'count'})
    group_orphans_main_df = group_orphans_main_df.join(brit_compens_eating_orphans_group_df)  # добавляем в свод
    group_orphans_main_df.rename(columns={'Статус_Сиротство': 'ПОО+компенсация'}, inplace=True)

    group_orphans_main_df.fillna(0, inplace=True)  # заполняем наны
    group_orphans_main_df = group_orphans_main_df.astype(int)  # приводим к инту
    sum_row = group_orphans_main_df.sum(axis=0)  # суммируем колонки
    group_orphans_main_df.loc['Итого'] = sum_row  # добавляем суммирующую колонку


    # отчет по учетам

    group_accounting_main_df = pd.DataFrame(index=list(df['Файл'].unique()))  # Базовый датафрейм для учетов

    # Считаем КДН
    kdn_accounting_df = df[df['Статус_КДН'].isin(['состоит'])]
    dct_name_sheet['КДН'] = kdn_accounting_df
    kdn_group_accounting_df = kdn_accounting_df.groupby(by=['Файл']).agg({'Статус_КДН': 'count'})
    group_accounting_main_df = group_accounting_main_df.join(kdn_group_accounting_df)  # добавляем в свод

    # Считаем ПДН
    pdn_accounting_df = df[df['Статус_ПДН'].isin(['состоит'])]
    dct_name_sheet['ПДН'] = pdn_accounting_df
    pdn_group_accounting_df = pdn_accounting_df.groupby(by=['Файл']).agg({'Статус_ПДН': 'count'})
    group_accounting_main_df = group_accounting_main_df.join(pdn_group_accounting_df)  # добавляем в свод

    # Слаживаем колонки и удаляем лишнее
    group_accounting_main_df.fillna(0, inplace=True)
    group_accounting_main_df['Учет КДН, ПДН'] = group_accounting_main_df['Статус_КДН'] + group_accounting_main_df[
        'Статус_ПДН']
    group_accounting_main_df.drop(columns=['Статус_КДН', 'Статус_ПДН'], inplace=True)

    # Считаем внутренний учет
    inside_accounting_df = df[df['Статус_Внутр_учет'].isin(['состоит'])]
    dct_name_sheet['Внутр_учет'] = inside_accounting_df
    inside_group_accounting_df = inside_accounting_df.groupby(by=['Файл']).agg({'Статус_Внутр_учет': 'count'})
    group_accounting_main_df = group_accounting_main_df.join(inside_group_accounting_df)  # добавляем в свод
    group_accounting_main_df.rename(columns={'Статус_Внутр_учет': 'Внутренний учет'}, inplace=True)

    # Считаем самовольный уход
    awol_accounting_df = df[df['Статус_Самовольный_уход'].isin(['да'])]
    dct_name_sheet['Самовольный уход'] = awol_accounting_df
    awol_group_accounting_df = awol_accounting_df.groupby(by=['Файл']).agg({'Статус_Самовольный_уход': 'count'})
    group_accounting_main_df = group_accounting_main_df.join(awol_group_accounting_df)  # добавляем в свод
    group_accounting_main_df.rename(columns={'Статус_Самовольный_уход': 'Самовольный уход'}, inplace=True)

    # Считаем самовольный уход
    sop_accounting_df = df[df['Статус_Соц_положение_семьи'].isin(['СОП'])]
    dct_name_sheet['СОП'] = sop_accounting_df
    sop_group_accounting_df = sop_accounting_df.groupby(by=['Файл']).agg({'Статус_Соц_положение_семьи': 'count'})
    group_accounting_main_df = group_accounting_main_df.join(sop_group_accounting_df)  # добавляем в свод
    group_accounting_main_df.rename(columns={'Статус_Соц_положение_семьи': 'СОП'}, inplace=True)

    group_accounting_main_df.fillna(0, inplace=True)  # заполняем наны
    group_accounting_main_df = group_accounting_main_df.astype(int)  # приводим к инту
    sum_row = group_accounting_main_df.sum(axis=0)  # суммируем колонки
    group_accounting_main_df.loc['Итого'] = sum_row  # добавляем суммирующую колонку

    # получаем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    # Создаем файл для самого отчета
    report_wb = write_df_to_excel_report_brit({'Социальный паспорт': group_main_df,'Сироты':group_orphans_main_df,'Учет':group_accounting_main_df}, write_index=True)
    report_wb = del_sheet(report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    report_wb.save(f'{path_end_folder}/Отчет по стандарту БРИТ от {current_time}.xlsx')

    # Сохраняем списки
    lst_report_wb = write_df_to_excel(dct_name_sheet, write_index=False)
    lst_report_wb = del_sheet(lst_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    lst_report_wb.save(f'{path_end_folder}/Списки для отчета по стандарту БРИТ от {current_time}.xlsx')


def create_social_report(etalon_file:str,data_folder:str,path_egisso_params:str, path_end_folder:str,checkbox_expelled:int,raw_date)->None:
    """
    Функция для генерации отчета по социальному статусу студентов БРИТ
    """
    try:
        set_rus_locale() # устанавливаем русскую локаль что категоризация по месяцам работала
        # обязательные колонки
        name_columns_set = {'ФИО','Дата_рождения','Статус_ОП','Статус_Бюджет','Статус_Общежитие','Статус_Учёба','Статус_Всеобуч', 'Статус_Национальность',
                            'Статус_Соц_стипендия', 'Статус_Соц_положение_семьи',
                            'Статус_Питание',
                            'Статус_Состав_семьи', 'Статус_Уровень_здоровья', 'Статус_Сиротство',
                            'Статус_Место_регистрации', 'Статус_Студенческая_семья',
                            'Статус_Воинский_учет','Статус_Родитель_СВО','Статус_Участник_СВО',
                            'Статус_ПДН','Статус_КДН','Статус_Внутр_учет','Статус_Спорт', 'Статус_Творчество',
                            'Статус_Волонтерство', 'Статус_Клуб', 'Статус_Самовольный_уход','Статус_Выпуск'}
        error_df = pd.DataFrame(
            columns=['Название файла', 'Название листа', 'Значение ошибки', 'Описание ошибки'])  # датафрейм для ошибок
        wb = openpyxl.load_workbook(etalon_file) # загружаем эталонный файл
        quantity_sheets = 0  # считаем количество групп
        main_sheet = wb.sheetnames[0] # получаем название первого листа с которым и будем сравнивать новые файлы
        main_df = pd.read_excel(etalon_file,sheet_name=main_sheet,nrows=0) # загружаем датафрейм чтобы получить эталонные колонки
        # Проверяем на обязательные колонки
        etallon_always_cols = name_columns_set.difference(set(main_df.columns))
        if len(etallon_always_cols) != 0:
            raise NotColumn
        etalon_cols = set(main_df.columns) # эталонные колонки

        for idx, file in enumerate(os.listdir(data_folder)):
            if not file.startswith('~$') and not file.endswith('.xlsx'):
                name_file = file.split('.xls')[0]
                temp_error_df = pd.DataFrame(data=[[f'{name_file}','', '',
                                                    'Расширение файла НЕ XLSX! Программа обрабатывает только XLSX ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                print(f'Файл: {name_file}')
                temp_wb = openpyxl.load_workbook(f'{data_folder}/{file}') # открываем
                lst_sheets_temp_wb = temp_wb.sheetnames # получаем список листов в файле
                for name_sheet in lst_sheets_temp_wb:
                    if name_sheet != 'Данные для выпадающих списков': # отбрасываем лист с даннными выпадающих списков
                        temp_df = pd.read_excel(f'{data_folder}/{file}',sheet_name=name_sheet,dtype=str) # получаем колонки которые есть на листе

                        # Проверяем на обязательные колонки
                        always_cols = name_columns_set.difference(set(temp_df.columns))
                        if len(always_cols) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}', f'{name_sheet}', f'{";".join(always_cols)}',
                                       'В файле на указанном листе не найдены указанные обязательные колонки. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                columns=['Название файла', 'Название листа', 'Значение ошибки',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                            continue  # не обрабатываем лист, где найдены ошибки

                        # проверяем на соответствие эталонному файлу
                        diff_cols = etalon_cols.difference(set(temp_df.columns))
                        if len(diff_cols) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}', f'{name_sheet}', f'{";".join(diff_cols)}',
                                       'В файле на указанном листе найдены лишние или отличающиеся колонки по сравнению с эталоном. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                columns=['Название файла', 'Название листа', 'Значение ошибки',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                            continue  # не обрабатываем лист, где найдены ошибки

                        # Проверяем на наличие колонок без названия Unnamed
                        unnamed_lst = [f'{idx} колонка не имеет названия' for idx, name_column in enumerate(temp_df.columns, 1) if 'Unnamed' in name_column]
                        if len(unnamed_lst) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}', f'{name_sheet}', f'{";".join(unnamed_lst)}',
                                       'В файле на указанном листе найдена(ы) колонка(и) у которых нет названия. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                columns=['Название файла', 'Название листа', 'Значение ошибки',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                            continue  # не обрабатываем лист, где найдены ошибки

                        # Проверяем порядок колонок
                        order_main_columns = list(main_df.columns)# порядок колонок эталонного файла
                        order_temp_df_columns = list(temp_df.columns)# порядок колонок проверяемого файла
                        error_order_lst = []# список для несовпадающих пар
                        # Сравниваем попарно колонки
                        for main,temp in zip(order_main_columns,order_temp_df_columns):
                            if main != temp:
                                error_order_lst.append(f'На месте колонки {main} находится колонка {temp}')
                        if len(error_order_lst) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}', f'{name_sheet}', f'{";".join(error_order_lst)}',
                                       'Неправильный порядок колонок по сравнению с эталоном. Приведите порядок колонок в соответствии с порядком в эталоне. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                columns=['Название файла', 'Название листа', 'Значение ошибки',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                            continue  # не обрабатываем лист, где найдены ошибки



                        temp_df.dropna(how='all', inplace=True)  # удаляем пустые строки
                        # проверяем наличие колонок Файл и Группа
                        if 'Файл' not in temp_df.columns:
                            temp_df.insert(0, 'Файл', name_file)

                        if 'Группа' not in temp_df.columns:
                            temp_df.insert(0, 'Группа', name_sheet) # вставляем колонку с именем листа

                        if checkbox_expelled == 0:
                            temp_df = temp_df[temp_df['Статус_Учёба'] != 'Отчислен'] # отбрасываем отчисленных если поставлен чекбокс

                        main_df = pd.concat([main_df,temp_df],axis=0,ignore_index=True) # добавляем в общий файл
                        quantity_sheets +=1
        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Проверяем есть ли данные в общем файле, если нет то вызываем исключение
        if len(main_df) == 0:
            error_df.to_excel(f'{path_end_folder}/Ошибки от {current_time}.xlsx',index=False)
            raise NotGoodSheet
        main_df.rename(columns={'Группа':'Для переноса','Файл':'файл для переноса'},inplace=True) # переименовываем группу чтобы перенести ее в начало таблицы
        main_df.insert(0,'Файл',main_df['файл для переноса'])
        main_df.insert(1,'Группа',main_df['Для переноса'])
        main_df.drop(columns=['Для переноса','файл для переноса'],inplace=True)

        main_df.fillna('Нет статуса', inplace=True) # заполняем пустые ячейки



        # Сохраняем лист со всеми данными
        lst_date_columns = []  # список для колонок с датами
        for column in main_df.columns:
            if 'дата' in column.lower():
                lst_date_columns.append(column)
        main_df[lst_date_columns] = main_df[lst_date_columns].applymap(convert_to_date)  # Приводим к типу
        main_df[lst_date_columns] = main_df[lst_date_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (pd.Timestamp,datetime.datetime)) and pd.notna(x) else x)

        main_df.replace('Нет статуса','',inplace=True)
        # Добавляем разбиение по датам
        main_df = proccessing_date(raw_date, 'Дата_рождения', main_df,path_end_folder)

        # Добавляем колонки со склоненными ФИО
        main_df = declension_fio_by_case(main_df)


        main_wb = write_df_to_excel({'Общий список':main_df},write_index=False)
        main_wb = del_sheet(main_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        main_wb.save(f'{path_end_folder}/Общий файл от {current_time}.xlsx')

        main_df.columns = list(map(str, list(main_df.columns)))


        # генерируем отчет по стандарту БРИТ
        create_report_brit(main_df.copy(),path_end_folder)


        # Генерируем файлы егиссо
        # генерируем полный вариант
        df_params_egisso,temp_params_egisso_error_df = extract_parameters_egisso(path_egisso_params,list(main_df.columns))
        if len(df_params_egisso) != 0:
            egisso_full_wb,egisso_not_find_wb, egisso_error_wb =create_full_egisso_data(main_df,df_params_egisso,path_end_folder) # создаем полный набор данных
            egisso_full_wb.save(f'{path_end_folder}/ЕГИССО полные данные от {current_time}.xlsx')
            egisso_not_find_wb.save(f'{path_end_folder}/ЕГИССО Не найденные льготы {current_time}.xlsx')
            egisso_error_wb.save(f'{path_end_folder}/ЕГИССО перс данные ОШИБКИ от {current_time}.xlsx')


        else:
            # генерируем вариант только с персональными данными
            egisso_clean_wb, egisso_error_wb = create_part_egisso_data(main_df)
            egisso_clean_wb.save(f'{path_end_folder}/ЕГИССО перс данные от {current_time}.xlsx')
            egisso_error_wb.save(f'{path_end_folder}/ЕГИССО перс данные ОШИБКИ от {current_time}.xlsx')

        # Сохраняем лист с ошибками

        error_df = pd.concat([error_df, temp_params_egisso_error_df], axis=0, ignore_index=True)
        error_wb = write_df_to_excel({'Ошибки':error_df},write_index=False)
        error_wb.save(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')

        # Создаем файл в котором будут сводные данные по колонкам с Подсчетом
        dct_counting_df = dict() # словарь в котором будут храниться датафреймы созданные для каждой колонки
        lst_counting_name_columns = [name_column for name_column in main_df.columns if 'Подсчет_' in name_column]
        if len(lst_counting_name_columns) != 0:
            for name_counting_column in lst_counting_name_columns:
                main_df[name_counting_column] = main_df[name_counting_column].apply(convert_number)
                temp_svod_df = (pd.pivot_table(main_df,index=['Файл','Группа'],
                                     values=[name_counting_column],
                                     aggfunc=[np.mean,np.sum,np.median,np.min,np.max,len]))
                temp_svod_df=temp_svod_df.reset_index() # убираем мультииндекс
                temp_svod_df = temp_svod_df.droplevel(axis=1,level=0) # убираем мультиколонки
                temp_svod_df.columns = ['Файл','Группа','Среднее','Сумма','Медиана','Минимум','Максимум','Количество']
                dct_counting_df[name_counting_column] = temp_svod_df # сохраняем в словарь

            # Сохраняем
            counting_report_wb = write_df_to_excel(dct_counting_df, write_index=False)
            counting_report_wb = del_sheet(counting_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
            counting_report_wb.save(f'{path_end_folder}/Свод по колонкам Подсчета от {current_time}.xlsx')

        # Создаем файл в котором будут данные по колонкам Список_
        dct_list_columns= {} # словарь в котором будут храниться датафреймы созданные для каждой колонки со списокм
        dct_values_in_list_columns = {} # словарь в котором будут храниться названия колонок и все значения которые там встречались
        dct_df_list_in_columns = {} # словарь где будут храниться значения в колонках и датафреймы где в указанных колонках есть соответствующее значение

        lst_list_name_columns = [name_column for name_column in main_df.columns if 'Список_' in name_column]
        if len(lst_list_name_columns) != 0:
            main_df[lst_list_name_columns] = main_df[lst_list_name_columns].astype(str)
            for name_lst_column in lst_list_name_columns:
                temp_col_value_lst = main_df[name_lst_column].tolist() # создаем список
                temp_col_value_lst = [value for value in temp_col_value_lst if value] # отбрасываем пустые значения
                unwrap_lst = []
                temp_col_value_lst = list(map(str,temp_col_value_lst)) # делаем строковым каждый элемент
                for value in temp_col_value_lst:
                    unwrap_lst.extend(value.split(','))
                unwrap_lst = list(map(str.strip,unwrap_lst)) # получаем список
                # убираем повторения и сортируем
                dct_values_in_list_columns[name_lst_column] = sorted(list(set(unwrap_lst)))

                dct_value_list = dict(Counter(unwrap_lst)) # Превращаем в словарь
                sorted_dct_value_lst = dict(sorted(dct_value_list.items())) # сортируем словарь
                # создаем датафрейм
                temp_df =  pd.DataFrame(list(sorted_dct_value_lst.items()), columns=['Показатель', 'Значение'])
                dct_list_columns[name_lst_column] = temp_df

            # Создаем датафреймы для подтверждения цифр
            for key,lst in dct_values_in_list_columns.items():
                for value in lst:
                    temp_list_df = main_df[main_df[key].str.contains(value)]
                    name_sheet = key.replace('Список_','')

                    dct_df_list_in_columns[f'{name_sheet}_{value}'] = temp_list_df

                # Сохраняем
            list_columns_report_wb = write_df_big_dct_to_excel(dct_df_list_in_columns, write_index=False)
            list_columns_report_wb = del_sheet(list_columns_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
            list_columns_report_wb.save(f'{path_end_folder}/Данные по своду Списков {current_time}.xlsx')


        # Создаем Свод по статусам
        main_df.replace('','Нет статуса',inplace=True)
        # Собираем колонки содержащие слово Статус_ и Подсчет_
        lst_status = [name_column for name_column in main_df.columns if 'Статус_' in name_column or 'Подсчет_' in name_column or 'Список_' in name_column]
        # Создаем датафрейм с данными по статусам
        soc_df = pd.DataFrame(columns=['Показатель','Значение']) # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп',quantity_sheets] # добавляем количество учебных групп
        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_Учёба'] == 'Обучается'].shape[0]  # со статусом Обучается
        quantity_academ_student = main_df[main_df['Статус_Учёба'].str.contains('Академический отпуск')].shape[
            0]
        quantity_not_status_student = main_df[main_df['Статус_Учёба'].str.contains('Нет статуса')].shape[
            0]
        quantity_except_deducted = main_df[~main_df['Статус_Учёба'].str.contains('Отчислен')].shape[
            0]  # все студенты кроме отчисленных
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',
                                   f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student}, Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)']  # добавляем количество студентов

        # считаем количество совершенолетних студентов
        quantity_maturity_students = len(main_df[main_df['Совершеннолетие'] == 'совершеннолетний'])
        quantity_not_maturity_students = len(main_df[main_df['Совершеннолетие'] == 'несовершеннолетний'])
        quantity_error_maturity_students = len(main_df[main_df['Совершеннолетие'].isin(['отрицательный возраст','Ошибочное значение!!!'])])
        soc_df.loc[len(soc_df)] = ['Возраст',
                                   f'Совершеннолетних - {quantity_maturity_students}, Несовершеннолетних - {quantity_not_maturity_students}, Неправильная дата рождения - {quantity_error_maturity_students}, Всего {quantity_except_deducted} (включая академ. и без статуса)']

        # Распределение по СПО-1
        header_spo = pd.DataFrame(columns=['Показатель', 'Значение'],
                                   data=[['Статус_СПО-1', None]])  # создаем строку с заголовком
        df_svod_by_SPO1 = main_df.groupby(['СПО_Один']).agg({'ФИО': 'count'})
        df_svod_by_SPO1 = df_svod_by_SPO1.reset_index()
        df_svod_by_SPO1.columns = ['Показатель', 'Значение']
        soc_df = pd.concat([soc_df, header_spo], axis=0)
        soc_df = pd.concat([soc_df, df_svod_by_SPO1], axis=0)
        for name_column in lst_status:
            if name_column == 'Статус_ОП':
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                # создаем строки с описанием
                new_value_df = create_value_str(main_df, name_column,'Статус_Учёба',
                                                {'Обучается': 'Обучается', 'Академ': 'Академический отпуск',
                                                 'Не указан статус': 'Нет статуса'})
            elif 'Статус_' in name_column:
                temp_counts = main_df[name_column].value_counts()  # делаем подсчет
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
                new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
                new_value_df['Показатель'] = new_value_df['Показатель'].astype(str)
                new_value_df.sort_values(by='Показатель',inplace=True)
            elif 'Подсчет_' in name_column:
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                main_df[name_column] = main_df[name_column].apply(convert_number)
                temp_desccribe = main_df[name_column].describe()
                sum_column = main_df[name_column].sum()
                _dct_describe = temp_desccribe.to_dict()
                dct_describe = {'Среднее':round(_dct_describe['mean'],2),'Сумма': round(sum_column,2),'Медиана':_dct_describe['50%'],
                                'Минимум':_dct_describe['min'],'Максимум':_dct_describe['max'],'Количество':_dct_describe['count'],}
                new_value_df = pd.DataFrame(list(dct_describe.items()),columns=['Показатель', 'Значение'])
            elif 'Список_' in name_column:
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                new_value_df = dct_list_columns[name_column]


            new_part_df = pd.concat([new_part_df, new_value_df], axis=0)  # соединяем
            soc_df = pd.concat([soc_df, new_part_df], axis=0)
        # Создаем раскладку по группам
        new_group_header_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                   data=[['Статус_студентов по группам', None]])  # создаем строку с заголовком
        new_group_df = create_value_str(main_df, 'Группа', 'Статус_Учёба',
                                        {'Обучается': 'Обучается', 'Академ': 'Академический отпуск',
                                         'Не указан статус': 'Нет статуса'})
        new_group_header_df = pd.concat([new_group_header_df, new_group_df], axis=0)

        soc_df = pd.concat([soc_df, new_group_header_df], axis=0)

        soc_wb = write_df_to_excel({'Свод по статусам':soc_df},write_index=False)
        soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])

        column_number = 0 # номер колонки в которой ищем слово Статус_
        # Создаем  стиль шрифта и заливки
        font = Font(color='FF000000')  # Черный цвет
        fill = PatternFill(fill_type='solid', fgColor='ffa500')  # Оранжевый цвет
        for row in soc_wb['Свод по статусам'].iter_rows(min_row=1, max_row=soc_wb['Свод по статусам'].max_row,
                                                        min_col=column_number, max_col=column_number):  # Перебираем строки
            if 'Статус_' in str(row[column_number].value) or 'Подсчет_' in str(row[column_number].value) or 'Список_' in str(row[column_number].value): # делаем ячейку строковой и проверяем наличие слова Статус_
                for cell in row: # применяем стиль если условие сработало
                    cell.font = font
                    cell.fill = fill
        soc_wb.save(f'{path_end_folder}/Свод по статусам от {current_time}.xlsx')

        # Создаем файл excel в котороым будет находится отчет
        wb = openpyxl.Workbook()

        # Проверяем наличие возможных дубликатов ,котороые могут получиться если обрезать по 30 символов
        lst_length_column = [column[:30] for column in main_df.columns]
        check_dupl_length = [k for k, v in Counter(lst_length_column).items() if v > 1]

        # проверяем наличие объединенных ячеек
        check_merge = [column for column in main_df.columns if 'Unnamed' in column]
        # если есть хоть один Unnamed то просто заменяем названия колонок на Колонка №цифра
        if check_merge or check_dupl_length:
            main_df.columns = [f'Колонка №{i}' for i in range(1, main_df.shape[1] + 1)]
        # очищаем названия колонок от символов */\ []''
        # Создаем регулярное выражение
        pattern_symbols = re.compile(r"[/*'\[\]/\\]")
        clean_main_df_columns = [re.sub(pattern_symbols, '', column) for column in main_df.columns]
        main_df.columns = clean_main_df_columns

        # Добавляем столбец для облегчения подсчета по категориям
        main_df['Для подсчета'] = 1

        # Создаем листы
        for idx, name_column in enumerate(main_df.columns):
            # Делаем короткое название не более 30 символов
            wb.create_sheet(title=name_column[:30], index=idx)

        for idx, name_column in enumerate(main_df.columns):
            group_main_df = main_df.astype({name_column:str}).groupby([name_column]).agg({'Для подсчета': 'sum'})
            group_main_df.columns = ['Количество']
            # Сортируем по убыванию
            group_main_df.sort_values(by=['Количество'], inplace=True, ascending=False)

            for r in dataframe_to_rows(group_main_df, index=True, header=True):
                if len(r) != 1:
                    wb[name_column[:30]].append(r)
            wb[name_column[:30]].column_dimensions['A'].width = 50

        # Удаляем листы
        wb = del_sheet(wb,['Sheet','Sheet1','Для подсчета'])
        # Сохраняем итоговый файл
        wb.save(f'{path_end_folder}/Свод по каждой колонке таблицы от {current_time}.xlsx')

        # проверяем на наличие ошибок
        if error_df.shape[0] != 0:
            count_error = len(error_df['Название листа'].unique())
            messagebox.showwarning('Деметра Отчеты социальный паспорт студента',
                                f'Количество необработанных листов {count_error}\n'
                                f'Проверьте файл Ошибки в файле')
    except FileNotFoundError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')
    except NotColumn:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Проверьте названия колонок в первом листе эталонного файла, для работы программы\n'
                             f' требуюется наличие колонок: {";".join(etallon_always_cols)}'
                             )
    except NotGoodSheet:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Заголовки ни одного листа не соответствуют эталонному файлу,\n'
                             f'Откройте файл с ошибками и устраните проблему'
                             )
    except ExceedingQuantity:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Количество групп или вариантов в колонках начинающихся с Список_ превышает 253 !\n'
                             f'Программа не может создать больше 253 листов в файле xlsx'
                             f'Сократите количество обрабатываемых значений')
    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Данные успешно обработаны')


if __name__ == '__main__':
    main_etalon_file = 'data/Эталон.xlsx'
    # main_etalon_file = 'data/Эталон подсчет.xlsx'
    main_data_folder = 'data/Данные'
    main_egisso_params = 'data/Параметры ЕГИССО.xlsx'
    main_end_folder = 'data/Результат'
    main_checkbox_expelled = 0
    main_raw_date = '05.09.2024'
    # main_checkbox_expelled = 1

    create_social_report(main_etalon_file,main_data_folder,main_egisso_params,main_end_folder,main_checkbox_expelled,main_raw_date)

    print('Lindy Booth !!!')

