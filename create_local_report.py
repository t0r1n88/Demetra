"""
Скрипт для обработки списка студентов на отделении и создания отчетности по нему
"""
from support_functions import *
import pandas as pd
import openpyxl
import time
from collections import Counter
import re
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

class NotStatusEdu(Exception):
    """
    Исключение для проверки наличия колонки Статус_учёба
    """
    pass



def prepare_file_params(params_file:str)->dict:
    """
    Функция для подготовки словаря с параметрами, преобразуюет первую колонку в ключи а вторую колонку в значения
    :param params_file: путь к файлу с параметрами в формате xlsx
    :return: словарь с параметрами
    """
    df =pd.read_excel(params_file,usecols='A:B',dtype=str)
    df.dropna(inplace=True) # удаляем все строки где есть нан
    lst_unique_name_column = df.iloc[:,0].unique() # получаем уникальные значения колонок в виде списка
    temp_dct = {key:{} for key in lst_unique_name_column} # создаем словарь верхнего уровня для хранения сгенерированных названий колонок
    # перебираем датафрейм и получаем
    for row in df.itertuples():
        # заполняем словарь
        temp_dct[row[1]][f'{row[1]}_{row[2]}'] = row[2]
    return temp_dct
def create_local_report(data_file_local:str, path_end_folder:str, params_report:str,checkbox_expelled:int) ->None:
    """
    Функция для генерации отчетов на основе файла с данными групп
    """
    try:
        main_df  = None # Базовый датафрейм , на основе первого лист
        error_df = pd.DataFrame(columns=['Лист','Ошибка','Примечание']) # датафрейм для ошибок
        example_columns = None # эталонные колонки
        temp_wb = openpyxl.load_workbook(data_file_local,read_only=True) # открываем файл для того чтобы узнать какие листы в нем есть
        lst_sheets = temp_wb.sheetnames
        lst_sheets = [name_sheet for name_sheet in lst_sheets if name_sheet != 'Данные для выпадающих списков']
        quantity_sheets = len(temp_wb.sheetnames)  # считаем количество групп
        temp_wb.close() # закрываем файл
        # словарь для основных параметров по которым нужно построить отчет
        dct_params = prepare_file_params(params_report) # получаем значения по которым нужно подсчитать данные
        lst_generate_name_columns = []  # создаем список для хранения значений сгенерированных колонок
        for key, value in dct_params.items():
            for name_gen_column in value.keys():
                lst_generate_name_columns.append(name_gen_column)
        custom_report_df = pd.DataFrame(columns=lst_generate_name_columns)
        custom_report_df.insert(0,'Лист',None)

        for name_sheet in lst_sheets:
            temp_df = pd.read_excel(data_file_local,sheet_name=name_sheet,dtype=str)
            temp_df.dropna(how='all',inplace=True) # удаляем пустые строки
            if '№ Группы' in temp_df.columns:
                temp_df.insert(0, '№ Группы_новый', name_sheet)  # вставляем колонку с именем листа
            else:
                temp_df.insert(0, '№ Группы', name_sheet)  # вставляем колонку с именем листа
            if not example_columns:
                if 'Статус_учёба' not in temp_df.columns:
                    raise NotStatusEdu
                example_columns = list(temp_df.columns) # делаем эталонным первый лист файла
                main_df = pd.DataFrame(columns=example_columns)
            # проверяем на соответствие колонкам первого листа
            if not set(example_columns).issubset(set(temp_df.columns)):
                diff_name_columns = set(example_columns).difference(set(temp_df.columns))
                error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_name_columns),
                                                                                       'Названия колонок указанного листа отличаются от названий колонок в первом листе. Исправьте отличия']])
                error_df = pd.concat([error_df,error_row],axis=0)
                continue
            if checkbox_expelled == 0:
                temp_df = temp_df[temp_df['Статус_учёба'] != 'Отчислен']
            main_df = pd.concat([main_df,temp_df],axis=0,ignore_index=True)

            # Подсчитываем основные показатели для каждой группы
            # проверяем наличие колонок в датафрейме
            diff_custom_name_columns = set(dct_params.keys()).difference(set(temp_df.columns))
            if len(diff_custom_name_columns) != 0:
                error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_custom_name_columns),
                                                                                       'Не найдены названия колонок в листе']])
                error_df = pd.concat([error_df,error_row],axis=0)
                continue

            row_dct = {key:0 for key in lst_generate_name_columns} # создаем словарь для хранения данных
            row_dct['Лист'] = name_sheet # добавляем колонки для листа
            for name_column,dct_value_column in dct_params.items():
                for key,value in dct_value_column.items():
                    row_dct[key] = temp_df[temp_df[name_column] == value].shape[0]
            new_row = pd.DataFrame(row_dct,index=[0])
            custom_report_df = pd.concat([custom_report_df,new_row],axis=0)
        # получаем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        # суммируем данные по листам
        all_custom_report_df = custom_report_df.sum(axis=0)
        all_custom_report_df = all_custom_report_df.drop('Лист').to_frame() # удаляем текстовую строку
        all_custom_report_df = all_custom_report_df.reset_index()
        all_custom_report_df.columns = ['Наименование параметра','Количество']
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


        # Удаляем листы
        wb = del_sheet(wb,['Sheet','Sheet1','Для подсчета'])
        # Сохраняем итоговый файл
        wb.save(f'{path_end_folder}/Отчет по всей таблице от {current_time}.xlsx')

        # Создаем Свод по статусам
        # Собираем колонки содержащие слово статус
        lst_status = [name_column for name_column in main_df.columns if 'Статус_' in name_column]

        status_df = main_df[lst_status] # оставляем датафрейм с данными статусов
        # Создаем датафрейм с данными по статусам
        soc_df = pd.DataFrame(columns=['Показатель','Значение']) # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп',quantity_sheets] # добавляем количество учебных групп
        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_учёба'] == 'Обучается'].shape[0]  # со статусом Обучается
        quantity_except_deducted = main_df[~main_df['Статус_учёба'].isin(['Нет статуса', 'Отчислен'])].shape[
            0]  # все студенты кроме отчисленных и у которых нет статуса
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',
                                   f'Обучается - {quantity_study_student}, Всего - {quantity_except_deducted}']  # добавляем количество студентов

        for name_column in lst_status:
            temp_counts = main_df[name_column].value_counts().sort_index()  # делаем подсчет
            new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                       data=[[name_column, None]])  # создаем строку с заголовком
            new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
            new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
            new_part_df = pd.concat([new_part_df, new_value_df], axis=0)  # соединяем
            soc_df = pd.concat([soc_df, new_part_df], axis=0)

        soc_wb = write_df_to_excel({'Свод по статусам':soc_df},write_index=False)
        soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        soc_wb.save(f'{path_end_folder}/Свод по статусам от {current_time}.xlsx')


        if error_df.shape[0] != 0:
            count_error = len(error_df['Лист'].unique())
            messagebox.showinfo('Деметра Отчеты социальный паспорт студента',
                                f'Количество необработанных листов {count_error}\n'
                                f'Проверьте файл Ошибки в файле')
    except FileNotFoundError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')
    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Данные успешно обработаны')


if __name__== '__main__':
    main_data_file = 'data/Пример файла.xlsx'
    main_result_folder = 'data/Результат'
    main_params_file = 'data/Параметры отчета.xlsx'
    main_checkbox_expelled = 0
    # main_checkbox_expelled = 1
    create_local_report(main_data_file, main_result_folder,main_params_file,main_checkbox_expelled)
    print('Lindy Booth')