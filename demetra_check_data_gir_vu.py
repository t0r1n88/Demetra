"""
Скрипт для массовой проверки и исправления файлов с данными ГИР ВУ
"""
import pandas as pd
import time
import os


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
                    # Проверяем порядок колонок
                    order_main_columns = lst_check_cols  # порядок колонок и названий как должно быть
                    order_temp_df_columns = list(df.columns)  # порядок колонок проверяемого файла
                    error_order_lst = []  # список для несовпадающих пар
                    # Сравниваем попарно колонки
                    for main, temp in zip(order_main_columns, order_temp_df_columns):
                        if main != temp:
                            error_order_lst.append(f'На месте колонки {main} находится колонка {temp}')
                    if len(error_order_lst) != 0:
                        error_order_message = ';'.join(error_order_lst)
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'{error_order_message}'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        count_errors += 1
                        continue

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









if __name__ == '__main__':
    main_data_folder = 'c:/Users/1/PycharmProjects/Demetra/data/ГИР ВУ'
    main_end_folder = 'c:/Users/1/PycharmProjects/Demetra/data/Результат'

    start_time = time.time()
    fixfiles_girvu(main_data_folder, main_end_folder)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Время выполнения: {elapsed_time:.6f} сек.")


    print('Lindy Booth')
