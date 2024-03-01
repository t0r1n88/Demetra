"""
Скрипт для создания  отчета по социальному паспорту студента БРИТ
"""
from support_functions import *
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
import time
from collections import Counter
import re
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
    new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                               data=[[name_column, None]])  # создаем строку с заголовком
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






def create_social_report(data_file_social:str, path_end_folder:str,checkbox_expelled:int)->None:
    """
    Функция для генерации отчета по социальному статусу студентов БРИТ
    """
    try:
        main_df  = None # Базовый датафрейм , на основе первого лист
        error_df = pd.DataFrame(columns=['Лист','Ошибка','Примечание']) # датафрейм для ошибок
        example_columns = None # эталонные колонки
        temp_wb = openpyxl.load_workbook(data_file_social,read_only=True) # открываем файл для того чтобы узнать какие листы в нем есть
        lst_sheets = temp_wb.sheetnames
        lst_sheets = [name_sheet for name_sheet in lst_sheets if name_sheet !='Данные для выпадающих списков']
        quantity_sheets = len(lst_sheets) # считаем количество групп
        temp_wb.close() # закрываем файл
        # обязательные колонки
        name_columns_set = {'Статус_ОП','Статус_Бюджет','Статус_Общежитие','Статус_Учёба','Статус_Всеобуч', 'Статус_Национальность', 'Статус_Соц_стипендия', 'Статус_Соц_положение_семьи',
                            'Статус_Питание',
                            'Статус_Состав_семьи', 'Статус_Уровень_здоровья', 'Статус_Сиротство',
                            'Статус_Отец_образование','Статус_Мать_образование','Статус_Опекун_Образование',
                            'Статус_Отец_сфера_деятельности','Статус_Мать_сфера_деятельности','Статус_Опекун_сфера_деятельности',
                            'Статус_Место_регистрации', 'Статус_Студенческая_семья',
                            'Статус_Воинский_учет','Статус_Родитель_СВО','Статус_Участник_СВО',
                            'Статус_ПДН','Статус_КДН','Статус_Нарк_учет','Статус_Внутр_учет','Статус_Спорт', 'Статус_Творчество',
                            'Статус_Волонтерство', 'Статус_Клуб', 'Статус_Самовольный_уход','Статус_Выпуск'}
        for name_sheet in lst_sheets:
            temp_df = pd.read_excel(data_file_social,sheet_name=name_sheet,dtype=str)
            temp_df.dropna(how='all',inplace=True) # удаляем пустые строки
            # проверяем наличие колонки № Группы
            if '№ Группы' in temp_df.columns:
                temp_df.insert(0, '№ Группы_новый', name_sheet)  # вставляем колонку с именем листа
            else:
                temp_df.insert(0, '№ Группы', name_sheet) # вставляем колонку с именем листа
            if not example_columns:
                example_columns = list(temp_df.columns) # делаем эталонным первый лист файла
                main_df = pd.DataFrame(columns=example_columns)
                # проверяем наличие колонок
                diff_first_sheet = name_columns_set.difference(set(temp_df.columns))
                if len(diff_first_sheet) != 0:
                    raise NotColumn

            # проверяем на соответствие колонкам первого листа
            if not set(example_columns).issubset(set(temp_df.columns)):
                diff_name_columns = set(example_columns).difference(set(temp_df.columns))
                error_row = pd.DataFrame(columns=['Лист','Ошибка','Примечание'],data=[[name_sheet,','.join(diff_name_columns),
                                                                                       'Названия колонок указанного листа отличаются от названий колонок в первом листе. Исправьте отличия']])
                error_df = pd.concat([error_df,error_row],axis=0)
                continue
            if checkbox_expelled == 0:
                temp_df = temp_df[temp_df['Статус_Учёба'] != 'Отчислен']
            main_df = pd.concat([main_df,temp_df],axis=0,ignore_index=True)

        main_df.fillna('Нет статуса', inplace=True) # заполняем пустые ячейки
        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

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
                                   f'Обучается - {quantity_study_student}, Академ - {quantity_academ_student}, Не указан статус - {quantity_not_status_student}, Всего {quantity_except_deducted} (включая академ. и без статуса)']  # добавляем количество студентов

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
        # проверяем на наличие ошибок
        if error_df.shape[0] != 0:
            count_error = len(error_df['Лист'].unique())
            messagebox.showinfo('Деметра Отчеты социальный паспорт студента',
                                f'Количество необработанных листов {count_error}\n'
                                f'Проверьте файл Ошибки в файле')
    except FileNotFoundError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')
    except NotColumn:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Проверьте названия колонок в первом листе файла с данными, для работы программы\n'
                             f' требуются колонки: {";".join(diff_first_sheet)}'
                             )
    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Данные успешно обработаны')


if __name__ == '__main__':
    main_data_file = 'data/Общий файл.xlsx'

    main_end_folder = 'data/Результат'
    main_checkbox_expelled = 0
    # main_checkbox_expelled = 1

    create_social_report(main_data_file,main_end_folder,main_checkbox_expelled)

    print('Lindy Booth !!!')

