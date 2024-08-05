"""
Функции для создания документов из шаблонов
"""

import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from docx2pdf import convert
from tkinter import messagebox
from jinja2 import exceptions
import time
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import platform
import logging
import tempfile
import re

logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)
class CheckBoxException(Exception):
    """
    Класс для вызовы исключения в случае если неправильно выставлены чекбоксы
    """
    pass


class NotFoundValue(Exception):
    """
    Класс для обозначения того что значение не найдено
    """
    pass

def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ValueError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    except TypeError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'


def processing_date_column(df, lst_columns):
    """
    Функция для обработки столбцов с датами. конвертация в строку формата ДД.ММ.ГГГГ
    """
    # получаем первую строку
    first_row = df.iloc[0, lst_columns]

    lst_first_row = list(first_row)  # Превращаем строку в список
    lst_date_columns = []  # Создаем список куда будем сохранять колонки в которых находятся даты
    tupl_row = list(zip(lst_columns,
                        lst_first_row))  # Создаем список кортежей формата (номер колонки,значение строки в этой колонке)

    for idx, value in tupl_row:  # Перебираем кортеж
        result = check_date_columns(idx, value)  # проверяем является ли значение датой
        if result:  # если да то добавляем список порядковый номер колонки
            lst_date_columns.append(result)
        else:  # иначе проверяем следующее значение
            continue
    for i in lst_date_columns:  # Перебираем список с колонками дат, превращаем их в даты и конвертируем в нужный строковый формат
        df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
        df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)

def check_date_columns(i, value):
    """
    Функция для проверки типа колонки. Необходимо найти колонки с датой
    :param i:
    :param value:
    :return:
    """
    try:
        itog = pd.to_datetime(str(value), infer_datetime_format=True)
    except:
        pass
    else:
        return i

def clean_value(value):
    """
    Функция для обработки значений колонки от  пустых пробелов,нан
    :param value: значение ячейки
    :return: очищенное значение
    """
    if value is np.nan:
        return 'Не заполнено'
    str_value = str(value)
    if str_value == '':
        return 'Не заполнено'
    elif str_value ==' ':
        return 'Не заполнено'

    return str_value



def combine_all_docx(filename_master, files_lst,mode_pdf,path_to_end_folder_doc,name_os):
    """
    Функция для объединения файлов Word взято отсюда
    https://stackoverflow.com/questions/24872527/combine-word-document-using-python-docx
    :param filename_master: базовый файл
    :param files_list: список с созданными файлами
    :return: итоговый файл
    """

    # Получаем текущее время
    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    number_of_sections = len(files_lst)
    # Открываем и обрабатываем базовый файл
    master = Document(filename_master)
    composer = Composer(master)
    # Перебираем и добавляем файлы к базовому
    for i in range(0, number_of_sections):
        doc_temp = Document(files_lst[i])
        composer.append(doc_temp)
    # Сохраняем файл
    composer.save(f"{path_to_end_folder_doc}/Объединеный файл от {current_time}.docx")
    if mode_pdf == 'Yes':
        if name_os == 'Windows':
            convert(f"{path_to_end_folder_doc}/Объединеный файл от {current_time}.docx",
                f"{path_to_end_folder_doc}/Объединеный файл от {current_time}.pdf", keep_active=True)
        else:
            raise NotImplementedError


def generate_docs_from_template(name_column,name_type_file,name_value_column,mode_pdf,name_file_template_doc,name_file_data_doc,path_to_end_folder_doc,
                                mode_combine,mode_group):
    """
    Функция для создания однотипных документов из шаблона Word и списка Excel
    :param name_column: название колонки в таблице данные из которой будут использоватьс для создания названий документов
    :param name_type_file: название создаваемых документов например Согласие,Справка и т.д.
    :param name_value_column: Значение из колонки name_type_file по которому будет создан единичный документ
    :param mode_pdf: чекбокс отвечающий за режим работы с pdf если Yes то будет создавать дополнительно pdf документ
    :param name_file_template_doc:путь к файлу шаблону на основе которого будут генерироваться документы
    :param name_file_data_doc: путь к файлу Excel с данными которые подставляются в шаблон
    :param path_to_end_folder_doc: путь к папке куда будут сохраняться файлы
    :param mode_combine:чекбокс отвечающий за режим объединения файлов. Если Yes то все документы будут объединены в один
    файл, если No то будет создаваться отдельный документ на каждую строчку исходной таблицы
    :param mode_group: чекбокс отвечающий за режим создания отдельного файла. Если Yes то можно создать один файл по значению
     из колонки name_value_column
    :return: Создает в зависимости от выбранного режима файлы Word из шаблона
    """
    try:
        name_os = platform.system() # получаем платформу на которой запущена программа
        # Считываем данные
        # Добавил параметр dtype =str чтобы данные не преобразовались а использовались так как в таблице
        df = pd.read_excel(name_file_data_doc, dtype=str)
        df[name_column] = df[name_column].apply(clean_value) # преобразовываем колонку меняя пустые значения и пустые пробелы на Не заполнено
        used_name_file = set()  # множество для уже использованных имен файлов
        # Заполняем Nan
        df.fillna(' ', inplace=True)
        lst_date_columns = []

        for idx, column in enumerate(df.columns):
            if 'дата' in column.lower():
                lst_date_columns.append(idx)

        # Конвертируем в пригодный строковый формат
        for i in lst_date_columns:
            df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
            df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')

        # В зависимости от состояния чекбоксов обрабатываем файлы
        if mode_combine == 'No':
            if mode_group == 'No':
                # Создаем в цикле документы
                for idx, row in enumerate(data):
                    doc = DocxTemplate(name_file_template_doc)
                    context = row
                    # print(context)
                    doc.render(context)
                    # Сохраняенм файл
                    # получаем название файла и убираем недопустимые символы < > : " /\ | ? *
                    name_file = row[name_column]
                    name_file = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_file)
                    # проверяем файл на наличие, если файл с таким названием уже существует то добавляем окончание
                    if name_file in used_name_file:
                        name_file = f'{name_file}_{idx}'

                    doc.save(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.docx')
                    used_name_file.add(name_file)
                    if mode_pdf == 'Yes':
                        if name_os == 'Windows':
                            convert(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.docx',
                                f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.pdf', keep_active=True)
                        else:
                            raise NotImplementedError
            else:
                # Отбираем по значению строку

                single_df = df[df[name_column] == name_value_column]
                # Конвертируем датафрейм в список словарей
                single_data = single_df.to_dict('records')
                # Проверяем количество найденных совпадений
                # очищаем от запрещенных символов
                name_file = name_value_column
                name_file = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_file)
                if len(single_data) == 1:
                    for row in single_data:
                        doc = DocxTemplate(name_file_template_doc)
                        doc.render(row)
                        # Сохраняенм файл
                        doc.save(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.docx')
                        if mode_pdf == 'Yes':
                            if name_os == 'Windows':
                                convert(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.docx',
                                    f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.pdf', keep_active=True)
                            else:
                                raise NotImplementedError

                elif len(single_data) > 1:
                    for idx, row in enumerate(single_data):
                        doc = DocxTemplate(name_file_template_doc)
                        doc.render(row)
                        # Сохраняем файл

                        doc.save(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}_{idx}.docx')
                        if mode_pdf == 'Yes':
                            if name_os == 'Windows':
                                convert(f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}_{idx}.docx',
                                    f'{path_to_end_folder_doc}/{name_type_file} {name_file[:80]}.pdf', keep_active=True)
                            else:
                                raise NotImplementedError
                else:
                    raise NotFoundValue



        else:
            if mode_group == 'No':
                # Список с созданными файлами
                files_lst = []
                # Создаем временную папку
                with tempfile.TemporaryDirectory() as tmpdirname:
                    print('created temporary directory', tmpdirname)
                    # Создаем и сохраняем во временную папку созданные документы Word
                    for idx,row in enumerate(data):
                        doc = DocxTemplate(name_file_template_doc)
                        context = row
                        doc.render(context)
                        # Сохраняем файл
                        # очищаем от запрещенных символов
                        name_file = f'{row[name_column]}'
                        name_file = re.sub(r'[\r\b\n\t<> :"?*|\\/]', '_', name_file)

                        doc.save(f'{tmpdirname}/{name_file[:80]}_{idx}.docx')
                        # Добавляем путь к файлу в список
                        files_lst.append(f'{tmpdirname}/{name_file[:80]}_{idx}.docx')
                    # Получаем базовый файл
                    main_doc = files_lst.pop(0)
                    # Запускаем функцию
                    combine_all_docx(main_doc, files_lst,mode_pdf,path_to_end_folder_doc,name_os)
            else:
                raise CheckBoxException

    except NameError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')
    except KeyError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'В таблице не найдена указанная колонка {e.args}')
    except PermissionError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Закройте все файлы Word созданные Вестой')
        logging.exception('AN ERROR HAS OCCURRED')
    except FileNotFoundError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Перенесите файлы, конечную папку с которой вы работете в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам или конечной папке.')
    except exceptions.TemplateSyntaxError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Ошибка в оформлении вставляемых значений в шаблоне\n'
                             f'Проверьте свой шаблон на наличие следующих ошибок:\n'
                             f'1) Вставляемые значения должны быть оформлены двойными фигурными скобками\n'
                             f'{{{{Вставляемое_значение}}}}\n'
                             f'2) В названии колонки в таблице откуда берутся данные - есть пробелы,цифры,знаки пунктуации и т.п.\n'
                             f'в названии колонки должны быть только буквы и нижнее подчеркивание.\n'
                             f'{{{{Дата_рождения}}}}')

    except NotImplementedError as e:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Создание pdf файлов работает ТОЛЬКО в WIndows, уберите галочку из чекбокса создания pdf ')
    except CheckBoxException:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Уберите галочку из чекбокса Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)'
                             )
    except NotFoundValue:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Указанное значение не найдено в выбранной колонке\nПроверьте наличие такого значения в таблице'
                             )
    except:
        logging.exception('AN ERROR HAS OCCURRED')
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')

    else:
        messagebox.showinfo('Веста Обработка таблиц и создание документов', 'Создание документов завершено!')

if __name__ == '__main__':
    name_column_main = 'ФИО'
    name_type_file_main = 'Справка'
    name_value_column_main = 'Алехин Данила Прокопьевич'
    mode_pdf_main = 'No'
    name_file_template_doc_main = 'data/Создание документов/Пример Шаблон согласия.docx'
    name_file_data_doc_main = 'data/Создание документов/Таблица для заполнения согласия.xlsx'
    path_to_end_folder_doc_main = 'data/result'
    mode_combine_main = 'No'
    mode_group_main = 'No'

    generate_docs_from_template(name_column_main, name_type_file_main, name_value_column_main, mode_pdf_main, name_file_template_doc_main,
                                name_file_data_doc_main, path_to_end_folder_doc_main,
                                mode_combine_main, mode_group_main)
    print('Lindy Booth')