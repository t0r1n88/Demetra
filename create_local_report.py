"""
Скрипт для обработки списка студентов на отделении и создания отчетности по нему
"""
from demetra_support_functions import write_df_to_excel, write_df_to_excel_report_brit, del_sheet, \
    declension_fio_by_case
from demetra_processing_date import proccessing_date
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


def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')


class NotColumn(Exception):
    """
    Исключение для обработки случая когда отсутствуют нужные колонки
    """
    pass


class NotCustomColumn(Exception):
    """
    Исключение для обработки случая когда отсутствуют колонки указанные в параметрах
    """
    pass


class NotGoodSheet(Exception):
    """
    Исключение для случая когда ни один лист не подхожит под эталон
    """
    pass


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


def create_value_str(df: pd.DataFrame, name_column: str, target_name_column: str, dct_str: dict) -> pd.DataFrame:
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
    new_value_df.sort_values(by='Показатель', inplace=True)
    for idx, row in enumerate(new_value_df.iterrows()):
        name_op = row[1].values[0]  # получаем название ОП
        temp_df = df[df[name_column] == name_op]  # отфильтровываем по названию ОП
        quantity_study_student = temp_df[temp_df[target_name_column] == dct_str['Обучается']].shape[
            0]  # со статусом Обучается
        quantity_academ_student = temp_df[temp_df[target_name_column].str.contains(dct_str['Академ'])].shape[
            0]
        quantity_not_status_student = \
        temp_df[temp_df[target_name_column].str.contains(dct_str['Не указан статус'])].shape[
            0]
        quantity_except_deducted = temp_df[~temp_df[target_name_column].str.contains('Отчислен')].shape[
            0]
        out_str = f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student},' \
                  f' Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)'
        new_value_df.iloc[idx, 1] = out_str  # присваиваем значение

    return new_value_df


def prepare_file_params(params_file: str):
    """
    Функция для подготовки словаря с параметрами, преобразуюет первую колонку в ключи а вторую колонку в значения
    :param params_file: путь к файлу с параметрами в формате xlsx
    :return: словарь с параметрами
    """
    df = pd.read_excel(params_file, usecols='A:B', dtype=str)
    df.dropna(inplace=True)  # удаляем все строки где есть нан
    lst_unique_name_column = df.iloc[:, 0].unique()  # получаем уникальные значения колонок в виде списка
    temp_dct = {key: {} for key in
                lst_unique_name_column}  # создаем словарь верхнего уровня для хранения сгенерированных названий колонок
    # перебираем датафрейм и получаем
    for row in df.itertuples():
        # заполняем словарь
        temp_dct[row[1]][f'{row[1]}_{row[2]}'] = row[2]
    return temp_dct, lst_unique_name_column, df


def create_for_custom_report(df: pd.DataFrame, params_df: pd.DataFrame) -> openpyxl.Workbook:
    """
    Функция для создания файла в котором в виде списков будут находиться данные использованные для создания настраиваемого отчета
    :param df: основной датафрейм
    :param params_df: датафрейм с параметрами
    :return:
    """
    used_name = set()  # множество для использованных названий листов
    dct_df = dict()  # словарь в котором будут храниться датафреймы
    for idx, row in enumerate(params_df.itertuples()):
        name_column = row[1]  # название колонки
        value_column = row[2]  # значение которое нужно подсчитать
        temp_df = df[df[name_column] == value_column]
        name_sheet = f'{name_column}'[:30]  # для того чтобы не было слишком длинного названия
        if name_sheet not in used_name:
            dct_df[name_sheet] = temp_df
            used_name.add(name_sheet)
        else:
            dct_df[f'{name_sheet}_{idx}'] = temp_df
            used_name.add(f'{name_sheet}_{idx}')

    lst_custom_wb = write_df_to_excel(dct_df, write_index=False)
    lst_custom_wb = del_sheet(lst_custom_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    return lst_custom_wb


def create_local_report(etalon_file: str, data_folder: str, path_end_folder: str, params_report: str,
                        checkbox_expelled: int, raw_date) -> None:
    """
    Функция для генерации отчетов на основе файла с данными групп
    """
    try:
        set_rus_locale()  # устанавливаем русскую локаль что категоризация по месяцам работала
        # обязательные колонки
        name_columns_set = {'Статус_ОП', 'Статус_Учёба', 'ФИО', 'Дата_рождения'}
        error_df = pd.DataFrame(
            columns=['Название файла', 'Название листа', 'Значение ошибки', 'Описание ошибки'])  # датафрейм для ошибок
        wb = openpyxl.load_workbook(etalon_file)  # загружаем эталонный файл
        quantity_sheets = 0  # считаем количество групп
        main_sheet = wb.sheetnames[0]  # получаем название первого листа с которым и будем сравнивать новые файлы
        main_df = pd.read_excel(etalon_file, sheet_name=main_sheet,
                                nrows=0)  # загружаем датафрейм чтобы получить эталонные колонки
        # Проверяем на обязательные колонки
        always_cols = name_columns_set.difference(set(main_df.columns))
        if len(always_cols) != 0:
            raise NotColumn
        etalon_cols = set(main_df.columns)  # эталонные колонки
        # словарь для основных параметров по которым нужно построить отчет
        # список уникальных названий колонок для проверки эталонного файла
        # датарфейм с параметрами, нужен для создания списков
        dct_params, lst_unique_params, params_df = prepare_file_params(
            params_report)  # получаем значения по которым нужно подсчитать данные и уникальные названия колонок
        # проверяем наличие колонок из файла параметров в эталонном файле
        custom_always_cols = set(lst_unique_params).difference(set(main_df.columns))
        if len(custom_always_cols) != 0:
            raise NotCustomColumn
        lst_generate_name_columns = []  # создаем список для хранения значений сгенерированных колонок
        for key, value in dct_params.items():
            for name_gen_column in value.keys():
                lst_generate_name_columns.append(name_gen_column)
        custom_report_df = pd.DataFrame(columns=lst_generate_name_columns)
        custom_report_df.insert(0, 'Файл', None)
        custom_report_df.insert(1, 'Лист', None)

        for idx, file in enumerate(os.listdir(data_folder)):
            if not file.startswith('~$') and not file.endswith('.xlsx'):
                name_file = file.split('.xls')[0]
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '', '',
                                                    'Расширение файла НЕ XLSX! Программа обрабатывает только XLSX ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                print(f'Файл: {name_file}')
                temp_wb = openpyxl.load_workbook(f'{data_folder}/{file}')  # открываем
                lst_sheets_temp_wb = temp_wb.sheetnames  # получаем список листов в файле
                for name_sheet in lst_sheets_temp_wb:
                    if name_sheet != 'Данные для выпадающих списков':  # отбрасываем лист с даннными выпадающих списков
                        temp_df = pd.read_excel(f'{data_folder}/{file}',
                                                sheet_name=name_sheet,dtype=str)  # получаем колонки которые есть на листе
                        # проверяем на соответствие эталонному файлу
                        diff_cols = etalon_cols.difference(set(temp_df.columns))
                        if len(diff_cols) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}', f'{name_sheet}', f'{";".join(diff_cols)}',
                                       'В файле на указанном листе найдены лишние или отличающиеся колонки по сравнению с эталоном. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                columns=['Название файла', 'Название листа', 'Значение ошибки',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                            continue  # не обрабатываем лист где найдены ошибки
                        if 'Файл' not in temp_df.columns:
                            temp_df.insert(0, 'Файл', name_file)

                        if 'Группа' not in temp_df.columns:
                            temp_df.insert(0, 'Группа', name_sheet)  # вставляем колонку с именем листа

                        if checkbox_expelled == 0:
                            temp_df = temp_df[
                                temp_df['Статус_Учёба'] != 'Отчислен']  # отбрасываем отчисленных если поставлен чекбокс

                        main_df = pd.concat([main_df, temp_df], axis=0, ignore_index=True)  # добавляем в общий файл
                        row_dct = {key: 0 for key in lst_generate_name_columns}  # создаем словарь для хранения данных
                        row_dct['Файл'] = name_file
                        row_dct['Лист'] = name_sheet  # добавляем колонки для листа
                        for name_column, dct_value_column in dct_params.items():
                            for key, value in dct_value_column.items():
                                row_dct[key] = temp_df[temp_df[name_column] == value].shape[0]
                        new_row = pd.DataFrame(row_dct, index=[0])
                        custom_report_df = pd.concat([custom_report_df, new_row], axis=0)

                        quantity_sheets += 1
        # получаем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        # Сохраняем лист с ошибками
        error_wb = write_df_to_excel({'Ошибки': error_df}, write_index=False)
        error_wb.save(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')
        if len(main_df) == 0:
            raise NotGoodSheet

        main_df.rename(columns={'Группа': 'Для переноса', 'Файл': 'файл для переноса'},
                       inplace=True)  # переименовываем группу чтобы перенести ее в начало таблицы
        main_df.insert(0, 'Файл', main_df['файл для переноса'])
        main_df.insert(1, 'Группа', main_df['Для переноса'])
        main_df.drop(columns=['Для переноса', 'файл для переноса'], inplace=True)
        main_df.fillna('Нет статуса', inplace=True)  # заполняем пустые ячейки

        # Сохраняем лист со всеми данными
        lst_date_columns = []  # список для колонок с датами
        for column in main_df.columns:
            if 'дата' in column.lower():
                lst_date_columns.append(column)
        main_df[lst_date_columns] = main_df[lst_date_columns].apply(pd.to_datetime, errors='coerce', dayfirst=True,
                                                                    format='mixed')  # Приводим к типу
        main_df[lst_date_columns] = main_df[lst_date_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (pd.Timestamp,datetime.datetime)) and pd.notna(x) else x)

        main_df.replace('Нет статуса', '', inplace=True)
        # Добавляем разбиение по датам
        main_df = proccessing_date(raw_date, 'Дата_рождения', main_df,path_end_folder)

        # Добавляем склонение по падежам и создание инициалов
        main_df = declension_fio_by_case(main_df)

        # Создаем списки на основе которых мы создаем настраиваемый отчет
        lst_custom_wb = create_for_custom_report(main_df, params_df)
        lst_custom_wb.save(f'{path_end_folder}/Списки для свода по выбранным колонкам от {current_time}.xlsx')

        # суммируем данные по листам

        all_custom_report_df = custom_report_df.sum(axis=0)
        all_custom_report_df = all_custom_report_df.drop(['Файл', 'Лист']).to_frame()  # удаляем текстовую строку

        all_custom_report_df = all_custom_report_df.reset_index()
        all_custom_report_df.columns = ['Наименование параметра', 'Количество']
        # сохраняем файл с данными по выбранным колонкам

        custom_report_wb = write_df_to_excel({'Общий свод': all_custom_report_df, 'Свод по листам': custom_report_df},
                                             write_index=False)
        custom_report_wb = del_sheet(custom_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        custom_report_wb.save(f'{path_end_folder}/Свод по выбранным колонкам Статусов от {current_time}.xlsx')

        # Сохраняем лист со всеми данными
        lst_date_columns = []  # список для колонок с датами
        for column in main_df.columns:
            if 'дата' in column.lower():
                lst_date_columns.append(column)
        main_df[lst_date_columns] = main_df[lst_date_columns].apply(pd.to_datetime, errors='coerce', dayfirst=True,
                                                                    format='mixed')  # Приводим к типу
        main_df[lst_date_columns] = main_df[lst_date_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, pd.Timestamp) else x)

        main_df.replace('Нет статуса', '', inplace=True)
        main_wb = write_df_to_excel({'Общий список': main_df}, write_index=False)
        main_wb = del_sheet(main_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        main_wb.save(f'{path_end_folder}/Общий файл от {current_time}.xlsx')

        main_df.columns = list(map(str, list(main_df.columns)))

        # Создаем файл в котором будут сводные данные по колонкам с Подсчетом
        dct_counting_df = dict()  # словарь в котором будут храниться датафреймы созданные для каждой колонки
        lst_counting_name_columns = [name_column for name_column in main_df.columns if 'Подсчет_' in name_column]
        if len(lst_counting_name_columns) != 0:
            for name_counting_column in lst_counting_name_columns:
                main_df[name_counting_column] = main_df[name_counting_column].apply(convert_number)
                temp_svod_df = (pd.pivot_table(main_df, index=['Файл', 'Группа'],
                                               values=[name_counting_column],
                                               aggfunc=[np.mean, np.sum, np.median, np.min, np.max, len]))
                temp_svod_df = temp_svod_df.reset_index()  # убираем мультииндекс
                temp_svod_df = temp_svod_df.droplevel(axis=1, level=0)  # убираем мультиколонки
                temp_svod_df.columns = ['Файл', 'Группа', 'Среднее', 'Сумма', 'Медиана', 'Минимум', 'Максимум',
                                        'Количество']
                dct_counting_df[name_counting_column] = temp_svod_df  # сохраняем в словарь

            # Сохраняем
            counting_report_wb = write_df_to_excel(dct_counting_df, write_index=False)
            counting_report_wb = del_sheet(counting_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
            counting_report_wb.save(f'{path_end_folder}/Свод по колонкам Подсчета от {current_time}.xlsx')

        # Создаем Свод по статусам
        # Собираем колонки содержащие слово Статус_ и Подсчет_
        lst_status = [name_column for name_column in main_df.columns if
                      'Статус_' in name_column or 'Подсчет_' in name_column]

        # Создаем датафрейм с данными по статусам
        soc_df = pd.DataFrame(columns=['Показатель', 'Значение'])  # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп', quantity_sheets]  # добавляем количество учебных групп
        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_Учёба'] == 'Обучается'].shape[0]  # со статусом Обучается
        quantity_academ_student = main_df[main_df['Статус_Учёба'].str.contains('Академический отпуск')].shape[
            0]
        quantity_not_status_student = main_df[main_df['Статус_Учёба'].str.contains('Нет статуса')].shape[
            0]
        quantity_except_deducted = main_df[~main_df['Статус_Учёба'].str.contains('Отчислен')].shape[
            0]  # все студенты кроме отчисленных
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',
                                   f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student},'
                                   f' Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)']  # добавляем количество студентов
        # считаем количество совершенолетних студентов
        quantity_maturity_students = len(main_df[main_df['Совершеннолетие'] == 'совершеннолетний'])
        quantity_not_maturity_students = len(main_df[main_df['Совершеннолетие'] == 'несовершеннолетний'])
        quantity_error_maturity_students = len(
            main_df[main_df['Совершеннолетие'].isin(['отрицательный возраст', 'Ошибочное значение!!!'])])
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
                new_value_df = create_value_str(main_df, name_column, 'Статус_Учёба',
                                                {'Обучается': 'Обучается', 'Академ': 'Академический отпуск',
                                                 'Не указан статус': 'Нет статуса'})
            elif 'Статус_' in name_column:
                temp_counts = main_df[name_column].value_counts()  # делаем подсчет
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
                new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
                new_value_df['Показатель'] = new_value_df['Показатель'].astype(str)
                new_value_df.sort_values(by='Показатель', inplace=True)
            elif 'Подсчет' in name_column:
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                main_df[name_column] = main_df[name_column].apply(convert_number)
                temp_desccribe = main_df[name_column].describe()
                sum_column = main_df[name_column].sum()
                _dct_describe = temp_desccribe.to_dict()
                dct_describe = {'Среднее': round(_dct_describe['mean'], 2), 'Сумма': round(sum_column, 2),
                                'Медиана': _dct_describe['50%'],
                                'Минимум': _dct_describe['min'], 'Максимум': _dct_describe['max'],
                                'Количество': _dct_describe['count'], }
                new_value_df = pd.DataFrame(list(dct_describe.items()), columns=['Показатель', 'Значение'])
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

        soc_wb = write_df_to_excel({'Свод по статусам': soc_df}, write_index=False)
        soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])

        column_number = 0  # номер колонки в которой ищем слово Статус_
        # Создаем  стиль шрифта и заливки
        font = Font(color='FF000000')  # Черный цвет
        fill = PatternFill(fill_type='solid', fgColor='ffa500')  # Оранжевый цвет
        for row in soc_wb['Свод по статусам'].iter_rows(min_row=1, max_row=soc_wb['Свод по статусам'].max_row,
                                                        min_col=column_number,
                                                        max_col=column_number):  # Перебираем строки
            if 'Статус_' in str(row[column_number].value) or 'Подсчет_' in str(
                    row[column_number].value):  # делаем ячейку строковой и проверяем наличие слова Статус_
                for cell in row:  # применяем стиль если условие сработало
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
            group_main_df = main_df.astype({name_column: str}).groupby([name_column]).agg({'Для подсчета': 'sum'})
            group_main_df.columns = ['Количество']

            # Сортируем по убыванию
            group_main_df.sort_values(by=['Количество'], inplace=True, ascending=False)

            for r in dataframe_to_rows(group_main_df, index=True, header=True):
                if len(r) != 1:
                    wb[name_column[:30]].append(r)
            wb[name_column[:30]].column_dimensions['A'].width = 50

        # Удаляем листы
        wb = del_sheet(wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        # Сохраняем итоговый файл
        wb.save(f'{path_end_folder}/Свод по каждой колонке таблицы от {current_time}.xlsx')

        if error_df.shape[0] != 0:
            count_error = len(error_df['Название листа'].unique())
            messagebox.showinfo('Деметра Отчеты социальный паспорт студента',
                                f'Количество необработанных листов {count_error}\n'
                                f'Проверьте файл Ошибки в файле')

    except FileNotFoundError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')
    except NotColumn:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Проверьте названия колонок в первом листе эталонного файла, для работы программы\n'
                             f' требуются колонки: {";".join(always_cols)}'
                             )
    except NotCustomColumn:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'В эталонном файле отсутствуют колонки указанные в файле с параметрами отчета,\n'
                             f' не найдены колонки: {";".join(custom_always_cols)}'
                             )
    except NotGoodSheet:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Заголовки ни одного листа не соответствуют эталонному файлу,\n'
                             f'Откройте файл с ошибками и устраните проблему'
                             )
    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Данные успешно обработаны')


if __name__ == '__main__':
    main_etalon_file = 'data/Эталон.xlsx'
    main_data_folder = 'data/Данные'
    main_result_folder = 'data/Результат'
    main_params_file = 'data/Параметры отчета.xlsx'
    main_checkbox_expelled = 0
    main_raw_data = '05.09.2024'
    # main_checkbox_expelled = 1
    create_local_report(main_etalon_file, main_data_folder, main_result_folder, main_params_file,
                        main_checkbox_expelled, main_raw_data)
    print('Lindy Booth')
