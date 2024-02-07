"""
Скрипт для создания  отчета по социальному паспорту студента БРИТ
"""
from support_functions import *
import pandas as pd
import openpyxl
import time


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

    """
    Вариатны подсчета показателей
    1 вариант это жестко заданный порядок, при этом не прописанные значения не попадают в список подсчета
    2 вариант это подсчет всего что есть, для всех колонок в названии котороых есть слово Статус
    """
    type_counting = 'not free'
    # type_counting = 'free'
    # Вариант с заданным порядком
    if type_counting == 'not free':
        soc_df = pd.DataFrame(columns=['Показатель','Значение']) # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп',quantity_sheets] # добавляем количество учебных групп


        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_учёба'] == 'Обучается'].shape[0] # со статусом Обучается
        quantity_except_deducted = main_df[~main_df['Статус_учёба'].isin(['Нет статуса','Отчислен'])].shape[0] # все студенты кроме отчисленных
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',f'{quantity_study_student}({quantity_except_deducted})'] # добавляем количество студентов

        # Создаем словарь для управления порядокм подсчета
        # dct_params = {'Статус_Национальность':['Русский','Бурят','КМН','Прочие','Нет статуса'],'Статус_Мат_положение':['Выше ПМ','Ниже ПМ']}
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
                      'Учет по профилактике': {'Статус_Профилактика': ['ПДН','КДН','наркологический учет','внутренний учет','СОП','несколько категорий','Нет статуса']},
                      'Охват студентов спортивными секциями и кружками': {'Статус_Спорт_доп': ['волейбол','баскетбол','борьба национальная','футбол','гиревой спорт','теннис','легкая атлетика','армспорт','шахматы','несколько категорий','Нет статуса']},
                      'Охват студентов творческими коллективами': {'Статус_Творчество_доп': ['вокал','танцы','ИЗО','театр','несколько категорий','Нет статуса']},
                      'Волонтерство': {'Статус_Волонтерство': ['да','Нет статуса',]},
                      'Клубы патриотические, военно-спортивные и другие': {'Статус_Клуб': ['патриотический','военно-спортивный','прочие','Нет статуса']}
                      }

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

    else:
        # Вариант со свободным порядком
        soc_df = pd.DataFrame(columns=['Показатель', 'Значение'])  # датафрейм для сбора данных отчета
        soc_df.loc[len(soc_df)] = ['Количество учебных групп', quantity_sheets]  # добавляем количество учебных групп

        # считаем количество студентов
        quantity_study_student = main_df[main_df['Статус_учёба'] == 'Обучается'].shape[0]  # со статусом Обучается
        quantity_except_deducted = main_df[~main_df['Статус_учёба'].isin(['Нет статуса', 'Отчислен'])].shape[
            0]  # все студенты кроме отчисленных
        soc_df.loc[len(soc_df)] = ['Количество студентов (контингент)',
                                   f'{quantity_study_student}({quantity_except_deducted})']  # добавляем количество студентов

        lst_status_column = [column for column in main_df.columns if 'Статус' in column] # получаем все колонки названия которых содержат слово Статус
        for name_column in lst_status_column:
            temp_counts = main_df[name_column].value_counts()  # делаем подсчет
            new_part_df = pd.DataFrame(columns=['Показатель', 'Значение'],
                                       data=[[name_column, None]])  # создаем строку с заголовком
            new_value_df = temp_counts.to_frame().reset_index()  # создаем датафрейм с данными
            new_value_df.columns = ['Показатель', 'Значение']  # делаем одинаковыми названия колонок
            new_part_df = pd.concat([new_part_df, new_value_df], axis=0)  # соединяем
            soc_df = pd.concat([soc_df, new_part_df], axis=0)

    # Сохраянем лист со всеми данными
    soc_wb = write_df_to_excel({'Социальный паспорт': soc_df}, write_index=False)
    soc_wb = del_sheet(soc_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
    soc_wb.save(f'{path_end_folder}/Социальный паспорт от {current_time}.xlsx')
    # Сохраняем лист с ошибками
    error_wb = write_df_to_excel({'Ошибки':error_df},write_index=False)
    error_wb.save(f'{path_end_folder}/Ошибки в файле от {current_time}.xlsx')




if __name__ == '__main__':
    main_data_file = 'data/Тестовая таблица ver 2.xlsx'
    main_end_folder = 'data/Результат'

    create_social_report(main_data_file,main_end_folder)

    print('Lindy Booth !!!')

