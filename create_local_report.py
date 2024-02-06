"""
Скрипт для обработки списка студентов на отделении и создания отчетности по нему
"""
from support_functions import *
import pandas as pd
import openpyxl
import time
from collections import Counter
import re

def prepare_file_params(params_file:str)->dict:
    """
    Функция для подготовки словаря с параметрами, преобразуюет первую колонку в ключи а вторую колонку в значения
    :param params_file: путь к файлу с параметрами в формате xlsx
    :return: словарь с параметрами
    """
    df =pd.read_excel(params_file,usecols='A:B',dtype=str)
    df.dropna(inplace=True) # удаляем все строки где есть нан
    temp_dct = dict(zip(df.iloc[:,0],df.iloc[:,1])) # создаем словарь с параметрами
    return temp_dct
def create_local_report(data_file_local:str, path_end_folder:str, params_report:str) ->None:
    """
    Функция для генерации отчетов на основе файла с данными групп
    """
    main_df  = None # Базовый датафрейм , на основе первого лист
    error_df = pd.DataFrame(columns=['Лист','Ошибка','Примечание']) # датафрейм для ошибок
    example_columns = None # эталонные колонки
    temp_wb = openpyxl.load_workbook(data_file_local,read_only=True) # открываем файл для того чтобы узнать какие листы в нем есть
    lst_sheets = temp_wb.sheetnames
    print(lst_sheets)
    temp_wb.close() # закрываем файл
    # словарь для основных параметров по которым нужно построить отчет
    dct_params = prepare_file_params(params_report) # получаем значения по которым нужно подсчитать данные
    lst_custom_name_columns = [f'{key}_{value}' for key,value in dct_params.items()] #
    custom_report_df = pd.DataFrame(columns=lst_custom_name_columns)
    custom_report_df.insert(0,'Лист',None)

    for name_sheet in lst_sheets:
        temp_df = pd.read_excel(data_file_local,sheet_name=name_sheet,dtype=str)
        temp_df.dropna(how='all',inplace=True) # удаляем пустые строки
        temp_df.insert(0, '№ Группы', name_sheet) # вставляем колонку с именем листа
        if not example_columns:
            example_columns = list(temp_df.columns) # делаем эталонным первый лист файла
            main_df = pd.DataFrame(columns=example_columns)
        # проверяем на соответствие колонкам первого листа
        diff_name_columns = set(temp_df.columns).difference(set(example_columns))
        if len(diff_name_columns) !=0:
            error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_name_columns),
                                                                                   'Названия колонок указанного листа отличаются от названий колонок в первом листе. Исправьте отличия']])
            error_df = pd.concat([error_df,error_row],axis=0)
            continue

        main_df = pd.concat([main_df,temp_df],axis=0,ignore_index=True)

        # Подсчитываем основные показатели для каждой группы
        # проверяем наличие колонок в датафрейме
        diff_custom_name_columns = set(dct_params.keys()).difference(set(temp_df.columns))
        if len(diff_custom_name_columns) != 0:
            error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_custom_name_columns),
                                                                                   'Не найдены названия колонок в листе']])
            error_df = pd.concat([error_df,error_row],axis=0)
            continue
        row_dct = {f'{key}_{value}':0 for key,value in dct_params.items()} # создаем словарь для создания строки датафрейма
        row_dct['Лист'] = name_sheet
        for key,value in dct_params.items():
            row_dct[f'{key}_{value}'] = temp_df[temp_df[key] == value].shape[0]
        new_row = pd.DataFrame(row_dct,index=[0])
        custom_report_df = pd.concat([custom_report_df,new_row],axis=0)
    # получаем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    # суммируем данные по листам
    all_custom_report_df = custom_report_df.sum(axis=0)
    all_custom_report_df = all_custom_report_df.drop('Лист').to_frame() # удаляем текстовую строку
    all_custom_report_df = all_custom_report_df.reset_index()
    all_custom_report_df.columns = ['Наименование параметра','Количство']
    # сохраняем файл с данными по выбранным колонкам

    custom_report_wb = write_df_to_excel({'Общий свод':all_custom_report_df,'Свод по листам':custom_report_df},write_index=False)
    custom_report_wb = del_sheet(custom_report_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    custom_report_wb.save(f'{path_end_folder}/Свод по выбранным колонкам от {current_time}.xlsx')

    # Сохраняем лист с ошибками
    error_wb = write_df_to_excel({'Ошибки':error_df},write_index=False)
    error_wb.save(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')
    # Сохраянем лист со всеми данными
    main_wb = write_df_to_excel({'Общий список':main_df},write_index=False)
    main_wb = del_sheet(main_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    main_wb.save(f'{path_end_folder}/Общий файл от {current_time}.xlsx')

    main_df.columns = list(map(str, list(main_df.columns)))

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
    # заполняем наны не заполнено
    main_df.fillna('Не заполнено',inplace=True)


    # Создаем листы
    for idx, name_column in enumerate(main_df.columns):
        # Делаем короткое название не более 30 символов
        wb.create_sheet(title=name_column[:30], index=idx)

    for idx, name_column in enumerate(main_df.columns):
        group_main_df = main_df.groupby([name_column]).agg({'Для подсчета': 'sum'})
        group_main_df.columns = ['Количество']

        # Сортируем по убыванию
        group_main_df.sort_values(by=['Количество'], inplace=True, ascending=False)

        for r in dataframe_to_rows(group_main_df, index=True, header=True):
            if len(r) != 1:
                wb[name_column[:30]].append(r)
        wb[name_column[:30]].column_dimensions['A'].width = 50

    # генерируем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    # Удаляем листы
    wb = del_sheet(wb,['Sheet','Sheet1','Для подсчета'])
    # Сохраняем итоговый файл
    wb.save(f'{path_end_folder}/Отчет по всей таблице от {current_time}.xlsx')

    main_df.to_excel('data/ghf.xlsx')




if __name__== '__main__':
    main_data_file = 'data/Тестовая таблица 1.xlsx'
    main_result_folder = 'data/Результат'
    main_params_file = 'data/Параметры отчета.xlsx'

    create_local_report(main_data_file, main_result_folder,main_params_file)
    print('Lindy Booth')