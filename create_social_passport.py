"""
Скрипт для создания  отчета по социальному паспорту студента БРИТ
"""
from support_functions import *
import pandas as pd
import openpyxl
import time
from collections import Counter
import re

def create_social_report(data_file_social:str, path_end_folder:str)->None:
    """
    Функция для генерации отчета по социальному статусу студентов БРИТ
    """
    main_df  = None # Базовый датафрейм , на основе первого лист
    error_df = pd.DataFrame(columns=['Лист','Ошибка','Примечание']) # датафрейм для ошибок
    example_columns = None # эталонные колонки
    temp_wb = openpyxl.load_workbook(data_file_social,read_only=True) # открываем файл для того чтобы узнать какие листы в нем есть
    lst_sheets = temp_wb.sheetnames
    quantity_sheets = len(temp_wb.sheetnames) # считаем количество групп
    temp_wb.close() # закрываем файл

    for name_sheet in lst_sheets:
        temp_df = pd.read_excel(data_file_social,sheet_name=name_sheet,dtype=str)
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

    main_df.fillna('Нет статуса', inplace=True) # заполняем пустые ячейки
    # генерируем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    # Сохраянем лист со всеми данными
    main_wb = write_df_to_excel({'Общий список':main_df},write_index=False)
    main_wb = del_sheet(main_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    main_wb.save(f'{path_end_folder}/Общий файл от {current_time}.xlsx')



    soc_df = pd.DataFrame(columns=['N','Показатель','Значение']) # датафрейм для сбора данных отчета
    soc_df.loc[len(soc_df)] = ['1','Количество учебных групп',quantity_sheets] # добавляем количество учебных групп


    # считаем количество студентов
    quantity_study_student = main_df[main_df['Статус_учёба'] == 'Обучается'].shape[0] # со статусом Обучается
    quantity_except_deducted = main_df[~main_df['Статус_учёба'].isin(['Нет статуса','Отчислен'])].shape[0] # все студенты кроме отчисленных
    soc_df.loc[len(soc_df)] = ['2','Количество студентов (контингент)',f'{quantity_study_student}({quantity_except_deducted})'] # добавляем количество студентов

    # Создаем словарь для управления порядокм подсчета
    dct_params = {'Статус_Национальность':['Русский','Бурят','КМН','Прочие','Нет статуса'],'Статус_Мат_положение':['Выше ПМ','Ниже ПМ']}

    # обрабатываем нужные колонки и упорядочиваем в правильном порядке
    for key,value in dct_params.items():
        pass








    print(soc_df)










if __name__ == '__main__':
    main_data_file = 'data/Тестовая таблица ver 2.xlsx'
    main_end_folder = 'data/Результат'

    create_social_report(main_data_file,main_end_folder)

    print('Lindy Booth !!!')

