"""
Скрипт для создания файла в котором будут содержаться частичные данные для загрузки в егиссо
Паспортные данные ,снилс фио
"""
from demetra_support_functions import write_to_excel_egisso
import pandas as pd
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

def find_cols_benefits(lst_cols:list):
    """
    Функция для поиска трех колонок отвечающих за определенные льготы
    :param lst_cols: список колонок датафрейма
    :return: словарь вида {Название колонки со статусом льготы: [Статус, Реквизиты, Дата истечения]}
    """
    # Проверяемые строки
    status_str = 'Статус_'
    requsit_str = 'Реквизит_'
    date_str = 'Дата_'

    # Словарь для хранения найденных колонок
    ben_dct = {}

    for i in range(len(lst_cols) - 2):  # Проходим по списку, оставляя место для двух следующих элементов
        if status_str in lst_cols[i] and  requsit_str in lst_cols[i + 1] and date_str in lst_cols[i + 2]:
            ben_dct[lst_cols[i]] = [lst_cols[i],lst_cols[i + 1],lst_cols[i + 2]]

    return ben_dct


def add_cols_pers_data(df:pd.DataFrame,ben_cols:list,req_cols_lst:list,name_col_ben:str):
    """
    Функция для добавления колонок с персональными данными
    :param df: полный датафрейм с данными
    :param ben_cols: список колонок относящихся к льготе
    :param req_cols_lst: наименования колонок которые нужно добавить
    :param name_col_ben: наименование льготы
    :return: датафрейм с добавленными колонками
    """
    df[name_col_ben] = df[name_col_ben].replace('нет',None) # подчищаем заполненные нет
    df = df[df[name_col_ben].notna()] # убираем незаполненные строки
    df = df[df[name_col_ben] != ''] # убираем пустые строки
    ben_df = df[ben_cols] # начинаем собирать датафрейм льгот
    ben_df.columns = ['Статус льготы','Реквизиты','Дата окончания льготы',]
    name_col_ben = name_col_ben.replace('Статус_','')
    ben_df.insert(0,'Льгота',name_col_ben) # добавляем колонку определяющую, что за льгота

    for name_column in req_cols_lst:
        if name_column in df.columns:
            ben_df[name_column] = df[name_column]
        else:
            ben_df[name_column] = f'В исходном файле не найдена колонка с названием {name_column}'
    ben_df.insert(10,'Тип документа','')

    return ben_df

def check_simple_str_column(value,error_str:str):
    """
    Функция для проверки на заполнение ячейки для простой колонки с текстом не требующим дополнительной проверки
    :param value: значение ячейки
    :param error_str: сообщение об ошибки
    """
    if pd.isna(value):
        return error_str
    else:
        return value

def processing_snils(value):
    """
    Функция для проверки и обработки СНИЛС
    :param value:
    :return:
    """
    result = re.findall(r'\d',value)
    if len(result) == 11:
        # проверяем на лидирующий ноль
        out_str = ''.join(result)
        if out_str.startswith('0'):
            return out_str
        else:
            return int(out_str)
    else:
        return f'Ошибка: В СНИЛС должно быть 11 цифр а в ячейке {len(result)} цифр(ы). В ячейке указано - {value}'

def processing_fio(value,pattern):
    """
    Функция для проверки соответствия
    :param value:значение
    :param pattern: объект re.compile
    :return:
    """
    if re.fullmatch(pattern,value):
        return value
    else:
        error_str = re.sub(r'\s','Пробельный символ',value)
        return f'Ошибка: ФИО должно начинаться с большой буквы и содержать только буквы кириллицы и дефис. В ячейке указано - {error_str}'

def comparison_date(value, pattern):
    """
    Функция для проверки соответсвия формата даты
    :param value:значение
    :param pattern: объект re.compile
    :return:
    """
    if re.fullmatch(pattern,value):
        return value
    else:
        return f'Ошибка: Дата должна иметь формат ДД.ММ.ГГГГ, например 21.10.2024. В ячейке указано - {value}'


def processing_series(value, pattern):
    """
    Функция для проверки соответсвия формата серии паспорта
    :param value:значение
    :param pattern: объект re.compile
    :return:
    """
    if re.fullmatch(pattern,value):
        if value.startswith('0'):
            return value
        else:
            return int(value)
    else:
        error_str = re.sub(r'\s','Пробельный символ',value)
        return f'Ошибка: Серия паспорта должна состоять из 4 цифр без пробелов, например 0343. В ячейке указано - {error_str}'

def processing_number(value, pattern):
    """
    Функция для проверки соответсвия формата номера паспорта
    :param value:значение
    :param pattern: объект re.compile
    :return:
    """

    if re.fullmatch(pattern,value):
        if value.startswith('0'):
            return value
        else:
            return int(value)
    else:
        error_str = re.sub(r'\s','Пробельный символ',value)
        return f'Ошибка: Номер паспорта должен состоять из 6 цифр без пробелов, например 420343. В ячейке указано - {error_str}'

def find_error_in_row(row):
    """
    Функция для поиска в каждой колонке строки слова Ошибка
    :param row:
    :return:
    """
    value_lst = row.tolist()
    error_lst = [value for value in value_lst if isinstance(value,str) and 'Ошибка' in value]
    if len(error_lst) !=0:
        return 'Ошибка'
    else:
        return 'Нет ошибок'
def check_error_ben(df:pd.DataFrame):
    """
    Функция для проверки правильности данных
    :param df:датафрейм с данными по одной льготе
    :return:2 датафрейма  один без ошибок и второй где указаны ошибки
    """

    checked_simple_cols = ['СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта']

    df[checked_simple_cols] = df[checked_simple_cols].applymap(lambda x:check_simple_str_column(x,'не заполнено'))
    df['СНИЛС'] = df['СНИЛС'].apply(processing_snils) # проверяем снилс и конвертируем снилс
    # првоеряем ФИО
    fio_pattern = re.compile(r'^[ЁА-Я][ёЁа-яА-Я-]+$')
    df['Фамилия'] = df['Фамилия'].apply(lambda x:processing_fio(x,fio_pattern)) # проверяем фамилию
    df['Имя'] = df['Имя'].apply(lambda x:processing_fio(x,fio_pattern)) # проверяем имя
    df['Отчество'] = df['Отчество'].apply(lambda x:processing_fio(x,fio_pattern)) # проверяем отчество

    # Проверяем М и Ж
    df['Пол'] = df['Пол'].apply(lambda x:x if x in ('М','Ж') else f'Ошибка: Допустимые значения М и Ж. В ячейке указано {x}')

    # проверяем колонку дату рождения
    date_pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$') # созадем паттерн
    df['Дата_рождения'] = df['Дата_рождения'].astype(str)
    df['Дата_рождения'] = df['Дата_рождения'].apply(lambda x:comparison_date(x, date_pattern))
    # Проверяем колонку серия паспорта
    series_pattern = re.compile(r'^\d{4}$')
    df['Серия_паспорта'] = df['Серия_паспорта'].astype(str)
    df['Серия_паспорта'] = df['Серия_паспорта'].apply(lambda x: processing_series(x, series_pattern))
    # проверяем номер паспорта
    number_pattern = re.compile(r'^\d{6}$')
    df['Номер_паспорта'] = df['Номер_паспорта'].astype(str)
    df['Номер_паспорта'] = df['Номер_паспорта'].apply(lambda x: processing_number(x, number_pattern))
    # проверяем колонку дата выдачи паспорта
    date_pattern = re.compile(r'^\d{2}\.\d{2}.\d{4}$') # созадем паттерн
    df['Дата_выдачи_паспорта'] = df['Дата_выдачи_паспорта'].astype(str)
    df['Дата_выдачи_паспорта'] = df['Дата_выдачи_паспорта'].apply(lambda x:comparison_date(x, date_pattern))
    # Проверяем колонку Кем выдано
    df['Кем_выдан'] = df['Кем_выдан'].apply(lambda x: check_simple_str_column(x, 'Ошибка: не заполнено'))
    df['Ошибка'] = df.apply(find_error_in_row,axis=1)

    # Создаем два датафрейма
    clean_df = df[df['Ошибка'] == 'Нет ошибок']
    error_df = df[df['Ошибка'] == 'Ошибка']
    # Убираем лишнюю колонку
    clean_df.drop(columns=['Ошибка'],inplace=True)
    error_df.drop(columns=['Ошибка'],inplace=True)

    return clean_df,error_df



def create_part_egisso_data(df:pd.DataFrame):
    """
    Функция для создания файла xlsx  содержащего в себе данные для егиссо
    ФИО,паспортные данные, снилс, колонки со льготами
    :param df: датафрейм с данными соц паспортов
    :return: 2 файла xlsx. С данными проверки корректности заполнения и с данными
    """
    main_df = pd.DataFrame(columns=['Льгота','Статус льготы','Реквизиты','Дата окончания льготы','Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан'])
    error_df = pd.DataFrame(columns=['Льгота','Статус льготы','Реквизиты','Дата окончания льготы','Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Тип документа','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан'])
    lst_cols_df = list(df.columns) # создаем список

    # ищем колонки со льготами
    benefits_cols_dct = find_cols_benefits(lst_cols_df)

    # список требуемых колонок для персональных данных
    req_lst_personal_data_cols = ['Файл','СНИЛС','Фамилия','Имя','Отчество','Пол','Дата_рождения','Серия_паспорта','Номер_паспорта',
                                  'Дата_выдачи_паспорта','Кем_выдан']

    # Собираем датафреймы
    for name_benefit,ben_cols in benefits_cols_dct.items():
        if name_benefit == 'Статус_Уровень_здоровья':
            health_df = df.copy() # костыль из-за того что в статус уровень здоровья для здоровых тоже указаны значения
            health_df['Статус_Уровень_здоровья'] = health_df['Статус_Уровень_здоровья'].fillna('доров')
            health_df = health_df[~health_df['Статус_Уровень_здоровья'].str.contains('доров')]
            temp_df_full = add_cols_pers_data(health_df,ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе
        else:
            temp_df_full = add_cols_pers_data(df.copy(),ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе

        temp_clean_df, temp_error_df =check_error_ben(temp_df_full)
        main_df = pd.concat([main_df,temp_clean_df])
        error_df = pd.concat([error_df,temp_error_df])


    main_wb = write_to_excel_egisso(main_df,'Чистый')
    error_wb = write_to_excel_egisso(error_df,'Ошибки')

    return main_wb,error_wb


def create_full_egisso_data(df:pd.DataFrame, params_egisso_df:pd.DataFrame):
    """
    Функция для создания полного файла ЕГИССО
    """
    main_df = pd.DataFrame(
        columns=['Льгота', 'Статус льготы', 'Реквизиты', 'Дата окончания льготы', 'Файл', 'СНИЛС', 'Фамилия', 'Имя',
                 'Отчество', 'Пол', 'Дата_рождения', 'Тип документа', 'Серия_паспорта', 'Номер_паспорта',
                 'Дата_выдачи_паспорта', 'Кем_выдан'])
    error_df = pd.DataFrame(
        columns=['Льгота', 'Статус льготы', 'Реквизиты', 'Дата окончания льготы', 'Файл', 'СНИЛС', 'Фамилия', 'Имя',
                 'Отчество', 'Пол', 'Дата_рождения', 'Тип документа', 'Серия_паспорта', 'Номер_паспорта',
                 'Дата_выдачи_паспорта', 'Кем_выдан'])
    lst_cols_df = list(df.columns)  # создаем список

    # ищем колонки со льготами
    benefits_cols_dct = find_cols_benefits(lst_cols_df)

    # список требуемых колонок для персональных данных
    req_lst_personal_data_cols = ['Файл', 'СНИЛС', 'Фамилия', 'Имя', 'Отчество', 'Пол', 'Дата_рождения',
                                  'Серия_паспорта', 'Номер_паспорта',
                                  'Дата_выдачи_паспорта', 'Кем_выдан']

    # Собираем датафреймы
    for name_benefit,ben_cols in benefits_cols_dct.items():
        if name_benefit == 'Статус_Уровень_здоровья':
            health_df = df.copy() # костыль из-за того что в статус уровень здоровья для здоровых тоже указаны значения
            health_df['Статус_Уровень_здоровья'] = health_df['Статус_Уровень_здоровья'].fillna('доров')
            health_df = health_df[~health_df['Статус_Уровень_здоровья'].str.contains('доров')]
            temp_df_full = add_cols_pers_data(health_df,ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе
        else:
            temp_df_full = add_cols_pers_data(df.copy(),ben_cols,req_lst_personal_data_cols,name_benefit) # получаем датафрейм по конкретной льготе

        temp_clean_df, temp_error_df =check_error_ben(temp_df_full)
        main_df = pd.concat([main_df,temp_clean_df])
        error_df = pd.concat([error_df,temp_error_df])


    # main_df['Соединение_ID'] = main_df['Льгота'] + main_df['Статус льготы']
    # params_egisso_df['Соединение_ID'] = params_egisso_df['Название колонки с льготой'] + params_egisso_df['Наименование категории']
    union_df = pd.merge(left=main_df,right=params_egisso_df,how='outer',left_on=['Льгота','Статус льготы'],
                       right_on=['Название колонки с льготой','Наименование категории'],indicator=True)

    clean_df = union_df[union_df['_merge'] == 'both'] # отбираем те льготы для котороых найдены совпадения.
    clean_df.drop(columns='_merge',inplace=True)
    not_find_ben_df = union_df[union_df['_merge'] != 'both']
    print(not_find_ben_df)

    # Возвращаем льготы и параметры льгот для котороых не нашли совпадения. ну и сам файл с данными.












if __name__ == '__main__':
    main_df = pd.read_excel('data/Свод.xlsx')

    create_part_egisso_data(main_df)


    print('Lindy Booth')
