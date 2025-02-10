"""
Скрипт для нахождения разницы между двумя таблицами
"""
from demetra_support_functions import write_df_to_excel, del_sheet # импорт функции по записи в файл с автошириной колонок
import pandas as pd
import re
import datetime
from tkinter import messagebox
import time
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None

# Классы для исключений
class ShapeDiffierence(Exception):
    """
    Класс для обозначения несовпадения размеров таблицы
    """
    pass


class ColumnsDifference(Exception):
    """
    Класс для обозначения того что названия колонок не совпадают
    """
    pass


def convert_to_date(value):
    """
    Функция для конвертации строки в текст
    :param value: значение для конвертации
    :return:
    """
    try:
        if value == 'Нет статуса':
            return 'Не заполнено'
        else:
            date_value = datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
            return date_value
    except ValueError:
        result = re.search(r'^\d{2}\.\d{2}\.\d{4}$', value)
        if result:
            try:
                return datetime.datetime.strptime(result.group(0), '%d.%m.%Y')
            except ValueError:
                # для случаев вида 45.09.2007
                return f'Некорректный формат даты - {value}, проверьте лишние пробелы,наличие точек'
        else:
            return f'Некорректный формат даты - {value}, проверьте лишние пробелы,наличие точек'
    except:
        return f'Некорректный формат даты - {value}, проверьте лишние пробелы,наличие точек'


def abs_diff(first_value, second_value):
    """
    Функция для подсчета абсолютной разницы между 2 значениями
    """
    try:
        return abs(float(first_value) - float(second_value))
    except:
        return None


def percent_diff(first_value, second_value):
    """
    функция для подсчета относительной разницы значений
    """
    try:
        # округляем до трех
        value = round(float(second_value) / float(first_value), 4) * 100
        return value
    except:
        return None


def change_perc_diff(first_value, second_value):
    """
    функция для подсчета процентного ихменения значений
    """
    try:
        value = (float(second_value) - float(first_value)) / float(first_value)
        return round(value, 4) * 100
    except:
        return None






def find_diffrence(first_df, second_df,path_to_end_folder_diffrence):
    """
    Функция для вычисления разницы между 2 таблицами
    :param first_df: Путь к первой таблице
    :param second_df: Путь ко второй таблице
    :param path_to_end_folder_diffrence : Путь к папке куда будут сохранятся файлы
    :return: разница между двумия таблица файл Excel в котором 2 листа:
    По колонкам - в котором указаны те ячейки в которых найдена разница
    По строкам - тоже самое только отображение по строкам
    """

    # загружаем датафреймы
    try:
        try:
            df1 = pd.read_excel(first_df, dtype=str)
            df2 = pd.read_excel(second_df, dtype=str)
        except:
            messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                                 f'Не удалось обработать файлы . Возможно какой то из файлов используемых для сравнения поврежден')

        # проверяем на соответсвие размеров
        if df1.shape != df2.shape:
            raise ShapeDiffierence

        # Проверям на соответсвие колонок
        if list(df1.columns) != list(df2.columns):
            diff_columns = set(df1.columns).difference(set(df2.columns))  # получаем отличающиеся элементы
            raise ColumnsDifference

        # Конвертируем даты в текстовый формат
        lst_date_columns = []  # список для колонок с датами
        for column in df1.columns:
            if 'дата' in column.lower():
                lst_date_columns.append(column)
        df1[lst_date_columns] = df1[lst_date_columns].applymap(convert_to_date)  # Приводим к типу
        df1[lst_date_columns] = df1[lst_date_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (pd.Timestamp, datetime.datetime)) and pd.notna(x) else x)

        lst_date_columns = []  # список для колонок с датами
        for column in df2.columns:
            if 'дата' in column.lower():
                lst_date_columns.append(column)
        df2[lst_date_columns] = df2[lst_date_columns].applymap(convert_to_date)  # Приводим к типу
        df2[lst_date_columns] = df2[lst_date_columns].applymap(
            lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (pd.Timestamp, datetime.datetime)) and pd.notna(x) else x)

        df_cols = df1.compare(df2,
                              result_names=('Первая таблица', 'Вторая таблица'))  # датафрейм с разницей по колонкам
        df_cols.index = list(
            map(lambda x: x + 2, df_cols.index))  # добавляем к индексу +2 чтобы соответствовать нумерации в экселе
        df_cols.index.name = '№ строки'  # переименовываем индекс

        df_rows = df1.compare(df2, align_axis=0,
                              result_names=('Первая таблица', 'Вторая таблица'))  # датафрейм с разницей по строкам
        lst_mul_ind = list(map(lambda x: (x[0] + 2, x[1]),
                               df_rows.index))  # добавляем к индексу +2 чтобы соответствовать нумерации в экселе
        index = pd.MultiIndex.from_tuples(lst_mul_ind, names=['№ строки', 'Таблица'])  # создаем мультиндекс
        df_rows.index = index

        # Создаем датафрейм с подсчетом разниц
        df_diff_cols = df_cols.copy()

        # получаем список колонок первого уровня
        temp_first_level_column = list(map(lambda x: x[0], df_diff_cols.columns))
        first_level_column = []
        [first_level_column.append(value) for value in temp_first_level_column if value not in first_level_column]

        # Добавляем колонки с абсолютной и относительной разницей
        count_columns = 2
        for name_column in first_level_column:
            # высчитываем абсолютную разницу
            df_diff_cols.insert(count_columns, (name_column, 'Разница между первым и вторым значением'),
                                df_diff_cols.apply(lambda x: abs_diff(x[name_column]['Первая таблица'],
                                                                      x[name_column]['Вторая таблица']), axis=1))

            # высчитываем отношение второго значения от первого
            df_diff_cols.insert(count_columns + 1, (name_column, '% второго от первого значения'),
                                df_diff_cols.apply(lambda x: percent_diff(x[name_column]['Первая таблица'],
                                                                          x[name_column]['Вторая таблица']), axis=1))

            # высчитываем процентное изменение
            df_diff_cols.insert(count_columns + 2, (name_column, 'Изменение в процентах'),
                                df_diff_cols.apply(lambda x: change_perc_diff(x[name_column]['Первая таблица'],
                                                                              x[name_column]['Вторая таблица']),
                                                   axis=1))

            count_columns += 5






        # записываем
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # записываем в файл Excel с сохранением ширины
        if len(df_cols) != 0:
            dct_df = {'По колонкам':df_cols,'По строкам':df_rows,'Значение разницы':df_diff_cols}
            write_index = True # нужно ли записывать индекс
            wb = write_df_to_excel(dct_df,write_index)
            wb = del_sheet(wb, ['Sheet', 'Sheet1'])
            wb.save(f'{path_to_end_folder_diffrence}/Разница между 2 таблицами {current_time}.xlsx')
        else:
            df_cols = pd.DataFrame(columns=['Отличий в таблицах не найдено'])
            dct_df = {'По колонкам':df_cols}
            write_index = True # нужно ли записывать индекс
            wb = write_df_to_excel(dct_df,write_index)
            wb = del_sheet(wb, ['Sheet', 'Sheet1'])
            wb.save(f'{path_to_end_folder_diffrence}/Разница между 2 таблицами {current_time}.xlsx')
    except UnboundLocalError:
        pass
    except ShapeDiffierence:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Не совпадают размеры таблиц, В первой таблице {df1.shape[0]}-стр. и {df1.shape[1]}-кол.\n'
                             f'Во второй таблице {df2.shape[0]}-стр. и {df2.shape[1]}-кол.')

    except ColumnsDifference:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Названия колонок в сравниваемых таблицах отличаются\n'
                             f'Колонок:{diff_columns}  нет во второй таблице !!!\n'
                             f'Сделайте названия колонок одинаковыми.')

    except FileNotFoundError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')

    else:
        messagebox.showinfo('Деметра Отчеты социальный паспорт студента', 'Таблицы успешно обработаны')



if __name__ == '__main__':
    main_df1 = 'data/1/БК-24 соцпедагога.xlsx'
    main_df2 = 'data/1/БК-24 куратора.xlsx'
    main_path_to_end_folder_diffrence = 'data/Результат'


    find_diffrence(main_df1,main_df2,main_path_to_end_folder_diffrence)

    print('Lindy Booth !!!')
