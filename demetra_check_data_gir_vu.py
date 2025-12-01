"""
Скрипт для массовой проверки и исправления файлов с данными ГИР ВУ
"""
from demetra_support_functions import (write_df_to_excel_cheking_egisso, del_sheet,write_df_error_egisso_to_excel,
                                       convert_to_date_gir_vu_cheking,create_doc_convert_date_egisso_cheking)
import pandas as pd
import time
import os
import re
from datetime import datetime

class NotFile(Exception):
    """
    Обработка случаев когда нет файлов в папке
    """
    pass

class BadOrderCols(Exception):
    """
    Исключение для обработки случая когда колонки не совпадают
    """
    pass


def strip_if_string(value):
    """Убирает пробелы вокруг строк"""
    if isinstance(value, str):
        return value.strip()
    return value

def delete_semicolon(value):
    """Убирает точку с запятой"""
    if isinstance(value, str):
        return value.replace(';',' ')
    return value


def drop_space_symbols(value:str):
    """
    Функция для замены пробельных символов в строке
    :param value:
    """
    if 'Ошибка' in value:
        return value

    result = re.sub(r'\s','',value)
    return result


def processing_fio(value):
    """
    Функция для обработки Фамилии или имени
    :param value:
    """
    if 'Ошибка' in value:
        return value

    value = re.sub(r'[^\s\w-]', '', value) # очищаем от всего кроме русских букв, пробела и тире
    value = re.sub(r'\s+', ' ', value)  # заменяем пробельные символы на один пробел

    # pattern = r'^[А-ЯЁа-яё]{0,30}(( |-)([А-ЯЁ][а-яё]{0,30})){0,2}$'
    pattern = r'^[А-ЯЁа-яё]{0,30}( |-)?([А-ЯЁа-яё]{0,30})?$'
    result = re.fullmatch(pattern, value)
    if result:
        if len(value) >100:
            return f'Ошибка: в значении {value} больше 100 символов'
        return f'{value[0].upper()}{value[1:].lower()}'
    else:
        return f'Ошибка: в значение {value}. Допустимы только буквы русского алфавита,дефис, пробел. Возможно лишний пробел рядом с дефисом или вместо русской буквы случайно записана английская. Например c-с или o-о'


def processing_patronymic(value):
    """
    Для обработки отчества
    :param value:
    """
    if isinstance(value, str):
        value = re.sub(r'[^\s\w-]', '', value)  # очищаем от всего кроме русских букв, пробела и тире
        value = re.sub(r'\s+', ' ', value)  # заменяем пробельные символы на один пробел


        # pattern = r'^[А-ЯЁа-яё]{0,30}(( |-)([А-ЯЁ][а-яё]{0,30})){0,2}$'
        pattern = r'^[А-ЯЁа-яё]{0,30}( |-)?([А-ЯЁа-яё]{0,30})?$'

        result = re.fullmatch(pattern, value)
        if result:
            if len(value) > 100:
                return f'Ошибка: в значении {value} больше 100 символов'
            return f'{value[0].upper()}{value[1:].lower()}'
        else:
            return f'Ошибка: в значение {value}. Допустимы только буквы русского алфавита,дефис, пробел. Возможно лишний пробел рядом с дефисом или вместо русской буквы случайно записана английская. Например c-с или o-о'

    else:
        return value

def processing_gender(value:str):
    """
    Функция для обработки колонки с полом
    :param value:
    """
    if 'Ошибка' in value:
        return value

    if value in ('0','1','2'):
        return int(value)

    if value[0].upper() == 'М':
        return 1
    elif value[0].upper() == 'Ж':
        return 2
    elif value.upper() == 'НЕ ОПРЕДЕЛЕНО':
        return 0
    else:
        return f'Ошибка: {value} неправильное значение'



def fixfiles_girvu(data_folder:str, end_folder:str):
    """
    Функция для проверки и исправления файлов ГИР ВУ
    :param data_folder: папка с файлами которые нужно проверить
    :param end_folder: конечная папка
    """
    count_errors = 0
    error_df = pd.DataFrame(
        columns=['Название файла', 'Описание ошибки'])  # датафрейм для ошибок

    lst_files = []  # список для файлов
    for dirpath, dirnames, filenames in os.walk(data_folder):
        lst_files.extend(filenames)
    # отбираем файлы
    lst_xlsx = [file for file in lst_files if not file.startswith('~$') and file.endswith('.xlsx')]
    quantity_files = len(lst_xlsx)  # считаем сколько xlsx файлов в папке

    # Обрабатываем в зависимости от количества файлов в папке
    if quantity_files == 0:
        raise NotFile
    else:
        lst_check_cols = ['Фамилия','Имя','Отчество',
                          'Пол (0-не определено, 1-мужской, 2-женский)', 'Дата рождения (ДД.ММ.ГГГГ.)',
                          'Серия паспорта гражданина РФ', 'Номер паспорта гражданина РФ', 'Дата выдачи паспорта гражданина РФ',
                          'СНИЛС гражданина (при наличии)', 'Наименование профессии, специальности, по которой проводится обучение (для программ СПО)',
                          'Код профессии, специальности, по которой проводится обучения (для программ СПО', 'Форма обучения', 'Номер курса',
                          'Полное наименование образовательной организации', 'Адрес образовательной организации', 'Дата поступления в образовательную организацию (ДД.ММ.ГГГГ)',
                          'Дата завершения обучения или отчисления из образовательной организации (ДД.ММ.ГГГГ.)'
                          ]

        # список колонок которые обязательно должны быть заполнены
        lst_required_filling = ['Фамилия','Имя','Отчество',
                          'Пол (0-не определено, 1-мужской, 2-женский)', 'Дата рождения (ДД.ММ.ГГГГ.)',
                          'Серия паспорта гражданина РФ', 'Номер паспорта гражданина РФ', 'Дата выдачи паспорта гражданина РФ',
                          'СНИЛС гражданина (при наличии)', 'Наименование профессии, специальности, по которой проводится обучение (для программ СПО)',
                          'Код профессии, специальности, по которой проводится обучения (для программ СПО', 'Форма обучения', 'Номер курса',
                          'Полное наименование образовательной организации', 'Адрес образовательной организации', 'Дата поступления в образовательную организацию (ДД.ММ.ГГГГ)',
                          'Дата завершения обучения или отчисления из образовательной организации (ДД.ММ.ГГГГ.)'
                                ]
        # lst_not_required_filling = [] # не требующие обязательного заполнения

        main_df = pd.DataFrame(columns=lst_check_cols)
        main_df.insert(0, 'Название файла', '')

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        for dirpath, dirnames, filenames in os.walk(data_folder):
            for file in filenames:
                if not file.startswith('~$') and file.endswith('.xlsx'):
                    try:
                        name_file = file.split('.xlsx')[0].strip()
                        print(name_file)  # обрабатываемый файл
                        df = pd.read_excel(f'{dirpath}/{file}', dtype=str)  # открываем файл
                    except:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'Не удалось обработать файл. Возможно файл поврежден'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        count_errors += 1
                        continue

                    # Проверяем на обязательные колонки
                    always_cols = set(lst_check_cols).difference(set(df.columns))
                    if len(always_cols) != 0:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}', f'{";".join(always_cols)}',
                                   'В файле на листе с данными не найдены указанные обязательные колонки. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                            columns=['Название файла', 'Значение ошибки',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                        continue  # не обрабатываем лист, где найдены ошибки

                    df = df[lst_check_cols]  # отбираем только обязательные колонки

                    if len(df) == 0:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'Файл пустой. Лист с данными должен быть первым по порядку'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        count_errors += 1
                        continue

                    # для строковых значений очищаем от пробельных символов в начале и конце
                    df = df.applymap(strip_if_string)
                    # очищаем от символа точка с запятой
                    df = df.applymap(delete_semicolon)

                    # Находим пропущенные значения в обязательных к заполнению колонках
                    df[lst_required_filling] = df[lst_required_filling].fillna('Ошибка: Ячейка не заполнена')
                    # Находим ячейки состоящие только из пробельных символов
                    # Регулярное выражение для поиска только пробелов
                    pattern_space = r'^[\s]*$'
                    # Заменяем ячейки, содержащие только пробельные символы, на нан
                    df[lst_required_filling] = df[lst_required_filling].replace(to_replace=pattern_space,
                                                                                value='Ошибка: Ячейка заполнена только пробельными символами',
                                                                                regex=True)

                    """
                    Начинаем проверять каждую колонку
                    """
                    df['Фамилия'] = df['Фамилия'].apply(processing_fio) # Фамилия
                    df['Имя'] = df['Имя'].apply(processing_fio) # Имя
                    df['Фамилия'] = df['Фамилия'].apply(processing_patronymic) # Отчество
                    # Пол
                    df['Пол (0-не определено, 1-мужской, 2-женский)'] = df['Пол (0-не определено, 1-мужской, 2-женский)'].apply(processing_gender)
                    # Дата рождения

                    current_date = datetime.now().date()  # Получаем текущую дату
                    # BirthDate_recip
                    df['Дата рождения (ДД.ММ.ГГГГ.)'] = df['Дата рождения (ДД.ММ.ГГГГ.)'].apply(
                        lambda x: convert_to_date_gir_vu_cheking(x, current_date))
                    df['Дата рождения (ДД.ММ.ГГГГ.)'] = df['Дата рождения (ДД.ММ.ГГГГ.)'].apply(create_doc_convert_date_egisso_cheking)






                    # Сохраняем датафрейм с ошибками разделенными по листам в соответсвии с колонками
                    dct_sheet_error_df = dict()  # создаем словарь для хранения названия и датафрейма

                    lst_name_columns = [name_cols for name_cols in df.columns if
                                        'Unnamed' not in name_cols]  # получаем список колонок

                    for idx, value in enumerate(lst_name_columns):
                        # получаем ошибки
                        temp_df = df[df[value].astype(str).str.contains('Ошибка')]  # фильтруем
                        if temp_df.shape[0] == 0:
                            continue

                        temp_df = temp_df[value].to_frame()  # оставляем только одну колонку

                        temp_df.insert(0, '№ строки с ошибкой в исходном файле',
                                       list(map(lambda x: x + 2, list(temp_df.index))))
                        dct_sheet_error_df[value] = temp_df

                    # создаем пути для проверки длины файла
                    error_path_file = f'{end_folder}/{name_file}/Базовые ошибки {name_file}.xlsx'
                    fix_path_file = f'{end_folder}/{name_file}/Обработанный {name_file}.xlsx'

                    # Сохраняем объединенные файлы
                    df.insert(0, 'Название файла', name_file)
                    main_df = pd.concat([main_df, df])



            main_error_wb = write_df_to_excel_cheking_egisso({'Критические ошибки':error_df},write_index=False)
            main_error_wb = del_sheet(main_error_wb,['Sheet', 'Sheet1', 'Для подсчета'])
            main_error_wb.save(f'{end_folder}/Критические ошибки {current_time}.xlsx')

            main_file_wb = write_df_error_egisso_to_excel({'Общий свод': main_df}, write_index=False)
            main_file_wb = del_sheet(main_file_wb, ['Sheet', 'Sheet1', 'Для подсчета'])
            main_file_wb.save(f'{end_folder}/Общий свод {current_time}.xlsx')











if __name__ == '__main__':
    main_data_folder = 'c:/Users/1/PycharmProjects/Demetra/data/ГИР ВУ'
    main_end_folder = 'c:/Users/1/PycharmProjects/Demetra/data/Результат'

    start_time = time.time()
    fixfiles_girvu(main_data_folder, main_end_folder)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Время выполнения: {elapsed_time:.6f} сек.")


    print('Lindy Booth')
