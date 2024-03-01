"""
Скрипт для обработки списка студентов на отделении и создания отчетности по нему
"""
from support_functions import *
import pandas as pd
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

class NotColumn(Exception):
    """
    Исключение для обработки случая когда отсутствуют нужные колонки
    """
    pass

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
def create_local_report(etalon_file:str,data_folder:str, path_end_folder:str, params_report:str,checkbox_expelled:int) ->None:
    """
    Функция для генерации отчетов на основе файла с данными групп
    """
    try:
        # обязательные колонки
        name_columns_set = {'Статус_ОП','Статус_Учёба'}
        error_df = pd.DataFrame(
            columns=['Название файла', 'Название листа', 'Значение ошибки', 'Описание ошибки'])  # датафрейм для ошибок
        wb = openpyxl.load_workbook(etalon_file) # загружаем эталонный файл
        quantity_sheets = 0  # считаем количество групп
        main_sheet = wb.sheetnames[0] # получаем название первого листа с которым и будем сравнивать новые файлы
        main_df = pd.read_excel(etalon_file,sheet_name=main_sheet,nrows=0) # загружаем датафрейм чтобы получить эталонные колонки
        # Проверяем на обязательные колонки
        always_cols = name_columns_set.difference(set(main_df.columns))
        if len(always_cols) != 0:
            raise NotColumn
        etalon_cols = set(main_df.columns) # эталонные колонки
        # словарь для основных параметров по которым нужно построить отчет
        dct_params = prepare_file_params(params_report) # получаем значения по которым нужно подсчитать данные
        lst_generate_name_columns = []  # создаем список для хранения значений сгенерированных колонок
        for key, value in dct_params.items():
            for name_gen_column in value.keys():
                lst_generate_name_columns.append(name_gen_column)
        custom_report_df = pd.DataFrame(columns=lst_generate_name_columns)
        custom_report_df.insert(0,'Файл',None)
        custom_report_df.insert(1,'Лист',None)

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
                        temp_df = pd.read_excel(f'{data_folder}/{file}',sheet_name=name_sheet) # получаем колонки которые есть на листе
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

                        if '№ Группы' in temp_df.columns:
                            temp_df.insert(0, '№ Группы_новый', name_sheet)  # вставляем колонку с именем листа
                        else:
                            temp_df.insert(0, '№ Группы', name_sheet) # вставляем колонку с именем листа

                        if checkbox_expelled == 0:
                            temp_df = temp_df[temp_df['Статус_Учёба'] != 'Отчислен'] # отбрасываем отчисленных если поставлен чекбокс

                        main_df = pd.concat([main_df,temp_df],axis=0,ignore_index=True) # добавляем в общий файл
                        row_dct = {key:0 for key in lst_generate_name_columns} # создаем словарь для хранения данных
                        row_dct['Файл'] =name_file
                        row_dct['Лист'] = name_sheet # добавляем колонки для листа
                        for name_column,dct_value_column in dct_params.items():
                            for key,value in dct_value_column.items():
                                row_dct[key] = temp_df[temp_df[name_column] == value].shape[0]
                        new_row = pd.DataFrame(row_dct,index=[0])
                        custom_report_df = pd.concat([custom_report_df,new_row],axis=0)



                        quantity_sheets += 1
        main_df.rename(columns={'№ Группы':'Для переноса','Файл':'файл для переноса'},inplace=True) # переименовываем группу чтобы перенести ее в начало таблицы
        main_df.insert(0,'Файл',main_df['файл для переноса'])
        main_df.insert(1,'№ Группы',main_df['Для переноса'])
        main_df.drop(columns=['Для переноса','файл для переноса'],inplace=True)

        main_df.fillna('Нет статуса', inplace=True) # заполняем пустые ячейки

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
        wb.save(f'{path_end_folder}/Свод по каждой колонке таблицы от {current_time}.xlsx')

        # Создаем Свод по статусам
        # Собираем колонки содержащие слово статус
        lst_status = [name_column for name_column in main_df.columns if 'Статус_' in name_column]

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
                                   f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student},'
                                   f' Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)']  # добавляем количество студентов
        for name_column in lst_status:
            if name_column == 'Статус_ОП':
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                # создаем строки с описанием
                new_value_df = create_value_str(main_df, name_column,'Статус_Учёба',
                                                {'Обучается': 'Обучается', 'Академ': 'Академический отпуск',
                                                 'Не указан статус': 'Нет статуса'})
            else:

                temp_counts = main_df[name_column].value_counts()  # делаем подсчет
                new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                           data=[[name_column, None]])  # создаем строку с заголовком
                new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
                new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
                new_value_df.sort_values(by='Показатель',inplace=True)
            new_part_df = pd.concat([new_part_df, new_value_df], axis=0)  # соединяем
            soc_df = pd.concat([soc_df, new_part_df], axis=0)
        # for name_column in lst_status:
        #     temp_counts = main_df[name_column].value_counts()  # делаем подсчет
        #     new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
        #                                data=[[name_column, None]])  # создаем строку с заголовком
        #     new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
        #     new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
        #     new_value_df.sort_values(by='Показатель',inplace=True)
        #     new_part_df = pd.concat([new_part_df, new_value_df], axis=0)  # соединяем
        #     soc_df = pd.concat([soc_df, new_part_df], axis=0)

        soc_wb = write_df_to_excel({'Свод по статусам':soc_df},write_index=False)
        soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])

        column_number = 0 # номер колонки в которой ищем слово Статус_
        # Создаем  стиль шрифта и заливки
        font = Font(color='FF000000')  # Черный цвет
        fill = PatternFill(fill_type='solid', fgColor='ffa500')  # Оранжевый цвет
        for row in soc_wb['Свод по статусам'].iter_rows(min_row=1, max_row=soc_wb['Свод по статусам'].max_row,
                                                        min_col=column_number, max_col=column_number):  # Перебираем строки
            if 'Статус_' in str(row[column_number].value): # делаем ячейку строковой и проверяем наличие слова Статус_
                for cell in row: # применяем стиль если условие сработало
                    cell.font = font
                    cell.fill = fill

        soc_wb.save(f'{path_end_folder}/Свод по статусам от {current_time}.xlsx')

        # Сохраняем лист с ошибками
        error_wb = write_df_to_excel({'Ошибки':error_df},write_index=False)
        error_wb.save(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')
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
    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Данные успешно обработаны')


if __name__== '__main__':
    main_etalon_file = 'data/Эталон.xlsx'
    main_data_folder = 'data/01.03'
    main_result_folder = 'data/Результат'
    main_params_file = 'data/Параметры отчета.xlsx'
    main_checkbox_expelled = 0
    # main_checkbox_expelled = 1
    create_local_report(main_etalon_file,main_data_folder, main_result_folder,main_params_file,main_checkbox_expelled)
    print('Lindy Booth')