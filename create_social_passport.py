"""
Скрипт для создания  отчета по социальному паспорту студента БРИТ
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
class NotColumn(Exception):
    """
    Исключение для обработки случая когда отсутствуют нужные колонки
    """
    pass


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
        quantity_sheets = len(temp_wb.sheetnames) # считаем количество групп
        temp_wb.close() # закрываем файл
        # обязательные колонки
        name_columns_set = {'Статус_общежитие','Статус_учёба', 'Статус_Национальность', 'Статус_Мат_положение', 'Статус_Соц_положение',
                            'Статус_Состав_семьи', 'Статус_Уровень_здоровья', 'Статус_Сиротство',
                            'Статус_Родители_образование',
                            'Статус_Родители_деятельность', 'Статус_Место_жительства', 'Статус_Студенческая_семья',
                            'Статус_ПДН','Статус_КДН','Статус_Нарк_учет','Статус_внутр_учет','Статус_Спорт_доп', 'Статус_Творчество_доп',
                            'Статус_Волонтерство', 'Статус_Клуб'}
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
                temp_df = temp_df[temp_df['Статус_учёба'] != 'Отчислен']
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
        wb.save(f'{path_end_folder}/Отчет по всей таблице от {current_time}.xlsx')


        soc_df = pd.DataFrame(columns=['Показатель','Значение']) # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп',quantity_sheets] # добавляем количество учебных групп


        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_учёба'] == 'Обучается'].shape[0] # со статусом Обучается
        quantity_except_deducted = main_df[~main_df['Статус_учёба'].isin(['Нет статуса','Отчислен'])].shape[0] # все студенты кроме отчисленных
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',f'{quantity_study_student}({quantity_except_deducted})'] # добавляем количество студентов

        # Создаем словарь для управления порядокм подсчета
        dct_params = {'Национальный состав':{'Статус_Национальность':['Русский','Бурят','КМН','Прочие','Нет статуса']},
                      'Материальное положение':{'Статус_Мат_положение':['Выше ПМ','Ниже ПМ','Нет статуса']},
                      'Социальное положение':{'Статус_Соц_положение':['Благополучное','СОП','Нет статуса']},
                      'Состав семьи': {'Статус_Состав_семьи': ['Полная','Неполная семья','Многодетная, не менее 3 несов.-летних. детей','Нет статуса']},
                      'Уровень здоровья': {'Статус_Уровень_здоровья': ['Здоров, нет противопоказаний','Имеют ограничения по здоровью (группа здоровья)','Инвалид','Нет статуса']},
                      'Дети-сироты и дети, оставшиеся без попечения родителей':{'Статус_Сиротство':['дети-сироты, находящиеся на полном государственном обеспечении','дети-сироты, находящиеся под опекой','Нет статуса']},
                      'Уровень образования родителей (законных представителей)':{'Статус_Родители_образование':['высшее','среднее профессиональное','среднее/основное общее образование','Нет статуса']},
                      'Сфера трудовой деятельности родителей (законных представителей)':{'Статус_Родители_деятельность':['бюджетная','коммерческая','правоохранительные структуры','пенсионеры','безработные','Нет статуса']},
                      'Место жительства':{'Статус_Место_жительства':['г.Улан-Удэ','муниципальное образование Республики Бурятия','регионы Российской Федерации','Нет статуса']},
                      'Студенческие семьи':{'Статус_Студенческая_семья':['оба родителя студенты','мать-одиночка, отец-одиночка','полная семья с детьми','Нет статуса']},
                      'Количество обучающихся состоящих на учете в ПДН':{'Статус_ПДН': ['состоит']},
                      'Количество обучающихся  состоящих на учете в КДН':{'Статус_КДН': ['состоит']},
                      'Количество обучающихся состоящих на наркологическом учете':{'Статус_Нарк_учет': ['состоит']},
                      'Количество обучающихся состоящих на внутреннем учете':{'Статус_внутр_учет': ['состоит']},
                      'Охват студентов спортивными секциями и кружками': {'Статус_Спорт_доп': ['волейбол','баскетбол','борьба национальная','футбол','гиревой спорт','теннис','легкая атлетика','армспорт','шахматы','несколько категорий','Нет статуса']},
                      'Охват студентов творческими коллективами': {'Статус_Творчество_доп': ['вокал','танцы','ИЗО','театр','несколько категорий','Нет статуса']},
                      'Волонтерство': {'Статус_Волонтерство': ['да','Нет статуса',]},
                      'Клубы патриотические, военно-спортивные и другие': {'Статус_Клуб': ['патриотический','военно-спортивный','прочие','Нет статуса']},
                      'Общежитие':{'Статус_общежитие':['да','Нет статуса']}}


        # обрабатываем нужные колонки и упорядочиваем в правильном порядке
        for indicator,dct_value in dct_params.items():
            for name_column,lst_order in dct_value.items():
                temp_counts = main_df[name_column].value_counts() # делаем подсчет
                temp_counts = temp_counts.reindex(lst_order) # меняем порядок на заданный
                new_part_df = pd.DataFrame(columns=['Показатель','Значение'],data=[[indicator,None]]) # создаем строку с заголовком
                new_value_df = temp_counts.to_frame().reset_index() # создаем датафрейм с данными
                new_value_df.columns = ['Показатель','Значение'] # делаем одинаковыми названия колонок
                new_part_df = pd.concat([new_part_df,new_value_df],axis=0) # соединяем
                soc_df=pd.concat([soc_df,new_part_df],axis=0)


        # Сохраянем лист со всеми данными
        soc_wb = write_df_to_excel({'Социальный паспорт': soc_df}, write_index=False)
        soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
        soc_wb.save(f'{path_end_folder}/Социальный паспорт от {current_time}.xlsx')
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
    main_data_file = 'data/Тестовая таблица ver 2.xlsx'
    main_data_file = 'data/Пример файла 05_02.xlsx'
    main_data_file = 'data/Пример файла.xlsx'
    main_end_folder = 'data/Результат'
    main_checkbox_expelled = 0
    # main_checkbox_expelled = 1

    create_social_report(main_data_file,main_end_folder,main_checkbox_expelled)

    print('Lindy Booth !!!')

