"""
Скрипт для обработки списка студентов на отделении и создания отчетности по нему
"""
import pandas as pd
import openpyxl
import time

def create_locac_report(data_file_local:str, path_end_folder:str) ->None:
    """
    Функция для генерации отчетов на основе файла с данными групп
    """
    main_df  = pd.DataFrame() # общий датафрейм добавлять туда будем по первому листу
    error_df = pd.DataFrame(columns=['Лист','Ошибка','Примечание'])
    example_columns = None
    temp_wb = openpyxl.load_workbook(data_file_local,read_only=True) # открываем файл для того чтобы узнать какие листы в нем есть
    lst_sheets = temp_wb.sheetnames
    print(lst_sheets)
    temp_wb.close() # закрываем файл
    for name_sheet in lst_sheets:
        temp_df = pd.read_excel(data_file_local,sheet_name=name_sheet,dtype=str)
        if not example_columns:
            example_columns = list(temp_df.columns)
        # проверяем на соответствие колонкам первого листа
        diff_name_columns = set(temp_df.columns).difference(set(example_columns))
        if len(diff_name_columns) !=0:
            error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_name_columns),
                                                                                   'Названия колонок указанного листа отличаются от названий колонок в первом листе. Исправьте отличия']])
            error_df = pd.concat([error_df,error_row],axis=0)
            print(diff_name_columns)
            continue

    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)
    error_df.to_excel(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')
















if __name__== '__main__':
    main_data_file = 'data/Данные.xlsx'
    main_result_folder = 'data'

    create_locac_report(main_data_file,main_result_folder)