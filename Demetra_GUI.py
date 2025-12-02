"""
Графический интерфейс для программы
"""
import re

from create_local_report import create_local_report  # создание отчета по выбранным пользователем параметрам
from create_social_passport import create_social_report  # создание отчета по социальному состоянию
from demetra_create_union_table import merge_table  # соединение таблиц
from demetra_checking_egisso import fix_files_egisso # Исправление файлов ЕГИССО
from demetra_check_data_gir_vu import fix_files_girvu
from expired_doc import check_expired_docs # поиск истекающих документов
from demetra_structuring_data import processing_structuring_data # Структурирование данных
from demetra_preparation_list import prepare_list  # подготовка персональных данных
from demetra_split_table import split_table  # разделение таблицы
from demetra_generate_docs import generate_docs_from_template # создание документов
from demetra_diff_tables import find_diffrence # нахождение разницы двух таблиц
from demetra_processing_birthday import proccessing_date # Функция для обработки дат рождения
from demetra_comparsion_two_tables import merging_two_tables # Функция для соединения 2 таблиц
import pandas as pd
from pandas._libs.tslibs.parsing import DateParseError
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import sys
import locale
import logging

logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)


class SameFolder(Exception):
    """
    Исключение для обработки случая когда выбраны одинаковые папки
    """
    pass


"""
Системные функции
"""

def set_rus_locale():
    """
    Функция чтобы можно было извлечь русские названия месяцев
    """
    locale.setlocale(
        locale.LC_ALL,
        'rus_rus' if sys.platform == 'win32' else 'ru_RU.UTF-8')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)


def on_scroll(*args):
    canvas.yview(*args)


def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.2)
    elif screen_width >= 2560:
        width = int(screen_width * 0.31)
    elif screen_width >= 1920:
        width = int(screen_width * 0.61)
    elif screen_width >= 1600:
        width = int(screen_width * 0.6)
    elif screen_width >= 1280:
        width = int(screen_width * 0.62)
    elif screen_width >= 1024:
        width = int(screen_width * 0.77)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")


"""
Прикладные функции
"""
"""
Создание локального отчета
"""


def select_file_etalon_local_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_etalon_local_report
    # Получаем путь к файлу
    name_file_etalon_local_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_file_params_local_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_params_local_report
    # Получаем путь к файлу
    name_file_params_local_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_folder_data_local_report():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_folder_local_report
    path_folder_local_report = filedialog.askdirectory()

def select_file_params_egisso_local_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_params_egisso_local_report
    # Получаем путь к файлу
    name_file_params_egisso_local_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_local_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_local_report
    path_to_end_folder_local_report = filedialog.askdirectory()


def processing_local_report():
    """
    Создание локальных отчетов
    :return:
    """
    try:
        select_date = var_select_date_local.get()
        # Если ничего
        if not select_date:
            select_date = pd.to_datetime('today')
        else:
            result = re.search(r'\d{2}\.\d{2}\.\d{4}', select_date)
            if result:
                select_date = pd.to_datetime(result.group(), dayfirst=True, errors='raise')
            else:
                raise DateParseError
        if path_folder_local_report == path_to_end_folder_local_report:
            raise SameFolder
        checkbox_expelled = group_rb_expelled_local_report.get()
        create_local_report(name_file_etalon_local_report, path_folder_local_report, path_to_end_folder_local_report,
                            name_file_params_local_report,name_file_params_egisso_local_report, checkbox_expelled,select_date)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите файл с параметрами,папку с данными, конечную папку')
    except SameFolder:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите разные папки в качестве исходной и конечной')
    except DateParseError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Введено некорректное значение даты.\n'
                             f'Введите дату в формате: ДД.ММ.ГГГГ например 14.06.2024')




"""
Создание социального отчета по контингенту БРИТ
"""


def select_file_etalon_social_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_etalon_social_report
    # Получаем путь к файлу
    name_file_etalon_social_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_folder_data_social_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_folder_social_report
    path_folder_social_report = filedialog.askdirectory()


def select_file_params_egisso_social_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_params_egisso_social_report
    # Получаем путь к файлу
    name_file_params_egisso_social_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))



def select_end_folder_social_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_social_report
    path_to_end_folder_social_report = filedialog.askdirectory()


def processing_social_report():
    """
    Создание отчета по социальным показателям БРИТ
    :return:
    """
    try:
        select_date = var_select_date.get()
        # Если ничего
        if not select_date:
            select_date = pd.to_datetime('today')
        else:
            result = re.search(r'\d{2}\.\d{2}\.\d{4}', select_date)
            if result:
                select_date = pd.to_datetime(result.group(), dayfirst=True, errors='raise')
            else:
                raise DateParseError
        if path_folder_social_report == path_to_end_folder_social_report:
            raise SameFolder
        checkbox_expelled = group_rb_expelled_social_report.get()
        create_social_report(name_file_etalon_social_report, path_folder_social_report,name_file_params_egisso_social_report,
                             path_to_end_folder_social_report, checkbox_expelled, select_date,)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите файл с параметрами,папку с данными, конечную папку')
    except SameFolder:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите разные папки в качестве исходной и конечной')
    except DateParseError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Введено некорректное значение даты.\n'
                             f'Введите дату в формате: ДД.ММ.ГГГГ например 14.06.2024')


"""
Функции для соединения таблиц
"""


def select_file_etalon_merge_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_etalon_merge_report
    # Получаем путь к файлу
    name_file_etalon_merge_report = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_data_folder_merge_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_data_folder_merge_report
    path_to_data_folder_merge_report = filedialog.askdirectory()


def select_end_folder_merge_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_merge_report
    path_to_end_folder_merge_report = filedialog.askdirectory()


def processing_merge_report():
    """
    Создание общей таблицы из нескольких файлов. При этом листы файлов копируются в общую таблицу сохраняя
    форматирование и данные
    :return: файл xlsx
    """
    try:
        merge_table(name_file_etalon_merge_report, path_to_data_folder_merge_report, path_to_end_folder_merge_report)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите файл с параметрами,файл с данными, конечную папку')


"""
Функции для разделения таблицы
"""


def select_file_split():
    """
    Функция для выбора файла с таблицей которую нужно разделить
    :return: Путь к файлу с данными
    """
    global file_data_split
    # Получаем путь к файлу
    file_data_split = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_split():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_split
    path_to_end_folder_split = filedialog.askdirectory()


def processing_split_table():
    """
    Функция для получения разделения таблицы по значениям
    :return:
    """
    # названия листов в таблицах
    try:
        # name_sheet = str(entry_sheet_name_split.get()) # получаем имя листа
        number_column = entry_number_column_split.get()  # получаем порядковый номер колонки
        number_column = int(number_column)  # конвертируем в инт

        checkbox_split = group_rb_type_split.get()  # получаем значения переключиталея

        # находим разницу
        split_table(file_data_split, number_column, checkbox_split, path_to_end_folder_split)
    except ValueError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Введите целое числа начиная с 1 !!!')
        logging.exception('AN ERROR HAS OCCURRED')
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для вкладки подготовка файлов
"""


def select_prep_file():
    """
    Функция для выбора файла который нужно преобразовать
    :return:
    """
    global glob_prep_file
    # Получаем путь к файлу
    glob_prep_file = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_prep():
    """
    Функция для выбора папки куда будет сохранен преобразованный файл
    :return:
    """
    global glob_path_to_end_folder_prep
    glob_path_to_end_folder_prep = filedialog.askdirectory()


def processing_preparation_file():
    """
    Функция для генерации документов
    """
    try:
        # name_sheet = var_name_sheet_prep.get() # получаем название листа
        checkbox_dupl = mode_dupl_value.get()
        checkbox_alf = mode_mix_alphabets.get()
        prepare_list(glob_prep_file, glob_path_to_end_folder_prep, checkbox_dupl,checkbox_alf)

    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для проверки истекающих документов
"""


def select_file_data_expired():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_data_expired
    # Получаем путь к файлу
    name_file_data_expired = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_data_folder_expired():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_data_folder_expired
    path_to_data_folder_expired = filedialog.askdirectory()


def processing_check_expired_docs():
    """
    Функция для генерации документов
    """
    try:
        check_expired_docs(name_file_data_expired, path_to_data_folder_expired)

    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')

"""
Функции структурирования данных
"""
def select_data_folder_extract_data_up():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_to_data_extract_data_up
    path_to_data_extract_data_up = filedialog.askdirectory()

def select_end_folder_extract_data_up():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_to_end_extract_data_up
    path_to_end_extract_data_up = filedialog.askdirectory()


def processing_extract_data_up():
    """
    Функция для структурирования данных
    """
    try:
        name_sheet = var_extract_name_sheet.get()
        quantity_header = var_extract_quantity_header.get()
        number_main_column = var_extract_number_main_column.get()
        quantity_cols = var_extract_quantity_cols.get()

        processing_structuring_data(path_to_data_extract_data_up, path_to_end_extract_data_up, name_sheet, quantity_header, number_main_column, quantity_cols)

    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')






"""
Функции для создания документов
"""


def select_file_template_doc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_doc
    name_file_template_doc = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_doc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_doc
    # Получаем путь к файлу
    name_file_data_doc = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_doc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_doc
    path_to_end_folder_doc = filedialog.askdirectory()


def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных Веста Обработка таблиц и создание документов ver 1.35)
    :return:
    """
    try:
        name_column = entry_name_column_data.get() # название колонки по которой будут создаваться имена файлов
        name_type_file = entry_type_file.get() # название создаваемого документа
        name_value_column = entry_value_column.get() # значение для генерации одиночного файла
        number_structure_folder = entry_structure_folder_value.get() # получаем список номеров колонок для структуры папок

        # получаем состояние чекбокса создания только pdf версий файлов
        mode_full = mode_full_value.get()
        # получаем состояние чекбокса создания pdf
        mode_pdf = mode_pdf_value.get()
        # Получаем состояние  чекбокса объединения файлов в один
        mode_combine = mode_combine_value.get()
        # Получаем состояние чекбокса создания индвидуального файла
        mode_group = mode_group_doc_value.get()
        # получаем состояние чекбокса создания структуры папок
        mode_structure_folder = mode_structure_folder_value.get()

        generate_docs_from_template(name_file_template_doc,name_file_data_doc,name_column, name_type_file, path_to_end_folder_doc, name_value_column, mode_pdf,
                                    mode_combine, mode_group,mode_structure_folder,number_structure_folder,mode_full)


    except NameError as e:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для разницы файлов
"""
def select_first_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_first_diffrence
    # Получаем путь к файлу
    data_first_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_second_diffrence
    # Получаем путь к файлу
    data_second_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_diffrence():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_diffrence
    path_to_end_folder_diffrence = filedialog.askdirectory()


def processing_diffrence():
    """
    Функция для нахождения разницы между двумя таблицами
    :return:
    """
    try:
        # находим разницу
        find_diffrence(data_first_diffrence, data_second_diffrence,
                       path_to_end_folder_diffrence)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')



"""
Функция для исправления файлов ЕГИССО
"""
def select_folder_data_fix_files_egisso():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_folder_fix_files_egisso
    path_folder_fix_files_egisso = filedialog.askdirectory()


def select_end_folder_data_fix_files_egisso():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_end_folder_fix_files_egisso
    path_end_folder_fix_files_egisso = filedialog.askdirectory()


def processing_fix_files_egisso():
    """
    Функция для подсчета
    :return:
    """
    try:
        if path_folder_fix_files_egisso == path_end_folder_fix_files_egisso:
            raise SameFolder
        # подсчитываем
        fix_files_egisso(path_folder_fix_files_egisso,path_end_folder_fix_files_egisso)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите папку с файлами и папку куда будет генерироваться результат')
    except SameFolder:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите разные папки в качестве исходной и конечной')


"""
Функции для исправления файлов ГИР ВУ
"""
def select_folder_data_fix_files_girvu():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_folder_fix_files_girvu
    path_folder_fix_files_girvu = filedialog.askdirectory()


def select_end_folder_data_fix_files_girvu():
    """
    Функция для выбора папки с данными
    :return:
    """
    global path_end_folder_fix_files_girvu
    path_end_folder_fix_files_girvu = filedialog.askdirectory()


def processing_fix_files_girvu():
    """
    Функция для подсчета
    :return:
    """
    try:
        if path_folder_fix_files_girvu == path_end_folder_fix_files_girvu:
            raise SameFolder
        # подсчитываем
        fix_files_girvu(path_folder_fix_files_girvu,path_end_folder_fix_files_girvu)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите папку с файлами и папку куда будет генерироваться результат')
    except SameFolder:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             'Выберите разные папки в качестве исходной и конечной')








"""
Функции для получения параметров обработки даты рождения
"""


def select_file_data_date():
    """
    Функция для выбора файла с данными для которого нужно разбить по категориям
    :return: Путь к файлу с данными
    """
    global name_file_data_date
    # Получаем путь к файлу
    name_file_data_date = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_date():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_date
    path_to_end_folder_date = filedialog.askdirectory()


def calculate_date():
    """
    Функция для разбиения по категориям, подсчета текущего возраста и выделения месяца,года
    :return:
    """
    try:
        raw_selected_date = entry_date.get()
        name_column = entry_name_column.get()
        # Устанавливаем русскую локаль
        set_rus_locale()
        proccessing_date(raw_selected_date, name_column, name_file_data_date, path_to_end_folder_date)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')



# Функциия для слияния 2 таблиц
def select_file_params_comparsion():
    """
    Функция для выбора файла с параметрами колонок т.е. кокие колонки нужно обрабатывать
    :return:
    """
    global file_params
    file_params = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_first_comparison():
    """
    Функция для выбора  первого файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_first_file_comparison
    # Получаем путь к файлу
    name_first_file_comparison = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_comparison():
    """
    Функция для выбора  второго файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_second_file_comparison
    # Получаем путь к файлу
    name_second_file_comparison = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_comparison():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_comparison
    path_to_end_folder_comparison = filedialog.askdirectory()


def processing_comparison():
    """
    Функция для сравнения,слияния 2 таблиц
    :return:
    """
    # получаем названия листов
    try:
        first_sheet = entry_first_sheet_name.get()
        second_sheet = entry_second_sheet_name.get()

        merging_two_tables(file_params, first_sheet, second_sheet, name_first_file_comparison,
                           name_second_file_comparison, path_to_end_folder_comparison)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')










"""
Создание нового окна
"""
def open_list_changes():
    # Создание нового окна
    new_window = Toplevel(window)

    # Настройка нового окна
    new_window.title("Список изменений")
    text_area = Text(new_window, width=90, height=50)

    with open(list_changes_path, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area.insert(END, text)
    text_area.configure(state='normal')
    text_area.pack(side=LEFT)

    scroll = Scrollbar(new_window, command=text_area.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area.config(yscrollcommand=scroll.set)

def open_license():
    # Создание нового окна
    new_window = Toplevel(window)

    # Настройка нового окна
    new_window.title("Лицензия")
    text_area = Text(new_window, width=90, height=50)

    with open(license_path, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area.insert(END, text)
    text_area.configure(state='normal')
    text_area.pack(side=LEFT)

    scroll = Scrollbar(new_window, command=text_area.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area.config(yscrollcommand=scroll.set)


def open_libraries():
    # Создание нового окна
    new_window = Toplevel(window)

    # Настройка нового окна
    new_window.title("Дополнительные библиотеки Python")
    text_area = Text(new_window, width=90, height=50)

    with open(license_library, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area.insert(END, text)
    text_area.configure(state='normal')
    text_area.pack(side=LEFT)

    scroll = Scrollbar(new_window, command=text_area.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area.config(yscrollcommand=scroll.set)




if __name__ == '__main__':
    window = Tk()
    window.title('Деметра Отчеты ver 2.31')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    # window.geometry('774x760')
    # window.geometry('980x910+700+100')
    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    global name_file_params_egisso_social_report
    name_file_params_egisso_social_report = 'Не выбрано' # костыль конечно но что поделать

    global name_file_params_egisso_local_report
    name_file_params_egisso_local_report = 'Не выбрано'  # костыль конечно но что поделать

    """
       Создаем вкладку для создания социального паспорта БРИТ
       """
    tab_create_social_report = ttk.Frame(tab_control)
    tab_control.add(tab_create_social_report, text='Стандартный\n отчет')

    create_social_report_frame_description = LabelFrame(tab_create_social_report)
    create_social_report_frame_description.pack()

    lbl_hello_create_social_report = Label(create_social_report_frame_description,
                                           text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                'Создание отчетов по стандарту БРИТ'
                                           , width=60)
    lbl_hello_create_social_report.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_social_report = resource_path('logo.png')
    img_create_social_report = PhotoImage(file=path_to_img_create_social_report)
    Label(create_social_report_frame_description,
          image=img_create_social_report, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того, чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_social_report = LabelFrame(tab_create_social_report, text='Подготовка')
    frame_data_social_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора эталонного файла
    btn_choose_file_etalon_social_report = Button(frame_data_social_report, text='1) Выберите эталонный файл',
                                                  font=('Arial Bold', 14),
                                                  command=select_file_etalon_social_report)
    btn_choose_file_etalon_social_report.pack(padx=10, pady=10)

    btn_choose_folder_social_report = Button(frame_data_social_report, text='2) Выберите папку с исходными файлами',
                                             font=('Arial Bold', 14),
                                             command=select_folder_data_social_report)
    btn_choose_folder_social_report.pack(padx=10, pady=10)

    # Определяем текстовую переменную для даты
    var_select_date = StringVar()
    # Описание поля
    label_select_date = Label(frame_data_social_report,
                              text='3) Введите дату на которую будет считаться текущий возраст студентов в формате: ДД.ММ.ГГГГ например 25.12.2024\n'
                                   'Если вы ничего не введете, то текущий возраст студентов будет считаться на момент запуска программы\n'
                                   'От значения текущего возраста зависит подсчет совершеннолетних, СПО-1, 1-ПК и т.п.')
    label_select_date.pack()
    # поле ввода
    entry_select_date = Entry(frame_data_social_report, textvariable=var_select_date, width=30)
    entry_select_date.pack()

    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_expelled_social_report = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_social_report = LabelFrame(frame_data_social_report, text='4) Выберите вариант подсчета')
    frame_rb_social_report.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_social_report, text='А) Подсчет без отчисленных', variable=group_rb_expelled_social_report,
                value=0).pack()
    Radiobutton(frame_rb_social_report, text='Б) Подсчет без студентов в академе и отчисленных', variable=group_rb_expelled_social_report,
                value=1).pack()
    Radiobutton(frame_rb_social_report, text='В) Подсчет с отчисленными и студентами в академе', variable=group_rb_expelled_social_report,
                value=2).pack()

    btn_choose_file_params_egisso_social_report = Button(frame_data_social_report, text='5) Выберите файл с параметрами ЕГИССО.\nНеобязательно, если вам не нужны файлы ЕГИССО',
                                                  font=('Arial Bold',8),
                                                  command=select_file_params_egisso_social_report)
    btn_choose_file_params_egisso_social_report.pack(padx=10, pady=10)


    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_social_report = Button(frame_data_social_report, text='6) Выберите конечную папку',
                                                 font=('Arial Bold', 14),
                                                 command=select_end_folder_social_report)
    btn_choose_end_folder_social_report.pack(padx=10, pady=10)

    # Создаем кнопку генерации отчетов

    btn_generate_social_report = Button(tab_create_social_report, text='7) Создать отчеты', font=('Arial Bold', 14),
                                        command=processing_social_report)
    btn_generate_social_report.pack(padx=10, pady=10)

    """
    Создаем вкладку для создания настраиваемых отчетов по любым таблицам
    """
    tab_create_local_report = ttk.Frame(tab_control)
    tab_control.add(tab_create_local_report, text='Настраиваемый\n отчет')

    create_local_report_frame_description = LabelFrame(tab_create_local_report)
    create_local_report_frame_description.pack()

    lbl_hello_create_local_report = Label(create_local_report_frame_description,
                                          text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                               'Создание настраиваемых отчетов'
                                          , width=60)
    lbl_hello_create_local_report.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_local_report = resource_path('logo.png')
    img_create_local_report = PhotoImage(file=path_to_img_create_local_report)
    Label(create_local_report_frame_description,
          image=img_create_local_report, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_local_report = LabelFrame(tab_create_local_report, text='Подготовка')
    frame_data_local_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора эталонного файла
    btn_choose_file_etalon_local_report = Button(frame_data_local_report, text='1) Выберите эталонный файл',
                                                 font=('Arial Bold', 14),
                                                 command=select_file_etalon_local_report)
    btn_choose_file_etalon_local_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с параметрами
    btn_choose_file_params_local_report = Button(frame_data_local_report, text='2) Выберите файл c параметрами',
                                                 font=('Arial Bold', 14),
                                                 command=select_file_params_local_report)
    btn_choose_file_params_local_report.pack(padx=10, pady=10)

    btn_choose_folder_data_local_report = Button(frame_data_local_report, text='3) Выберите папку с исходными файлами',
                                                 font=('Arial Bold', 14),
                                                 command=select_folder_data_local_report)
    btn_choose_folder_data_local_report.pack(padx=10, pady=10)
    # Определяем текстовую переменную для даты
    var_select_date_local = StringVar()
    # Описание поля
    label_select_date_local = Label(frame_data_local_report,
                              text='4) Введите дату на которую будет считаться текущий возраст студентов в формате: ДД.ММ.ГГГГ например 25.12.2024\n'
                                   'Если вы ничего не введете, то текущий возраст студентов будет считаться на момент запуска программы\n'
                                   'От значения текущего возраста зависит подсчет совершеннолетних, СПО-1, 1-ПК и т.п.')
    label_select_date_local.pack()
    # поле ввода
    entry_select_date_local = Entry(frame_data_local_report, textvariable=var_select_date_local, width=30)
    entry_select_date_local.pack()

    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_expelled_local_report = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_local_report = LabelFrame(frame_data_local_report, text='5) Выберите вариант подсчета')
    frame_rb_local_report.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_local_report, text='А) Подсчет без отчисленных', variable=group_rb_expelled_local_report,
                value=0).pack()
    Radiobutton(frame_rb_local_report, text='Б) Подсчет без студентов в академе и отчисленных', variable=group_rb_expelled_local_report,
                value=1).pack()
    Radiobutton(frame_rb_local_report, text='В) Подсчет с отчисленными и студентами в академе', variable=group_rb_expelled_local_report,
                value=2).pack()

    btn_choose_file_params_egisso_local_report = Button(frame_data_local_report, text='6) Выберите файл с параметрами ЕГИССО.\nНеобязательно, если вам не нужны файлы ЕГИССО',
                                                  font=('Arial Bold', 8),
                                                  command=select_file_params_egisso_local_report)
    btn_choose_file_params_egisso_local_report.pack(padx=10, pady=10)


    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_local_report = Button(frame_data_local_report, text='7) Выберите конечную папку',
                                                font=('Arial Bold', 14),
                                                command=select_end_folder_local_report)
    btn_choose_end_folder_local_report.pack(padx=10, pady=10)

    # Создаем кнопку генерации отчетов

    btn_generate_local_report = Button(tab_create_local_report, text='8) Создать отчеты', font=('Arial Bold', 14),
                                       command=processing_local_report)
    btn_generate_local_report.pack(padx=10, pady=10)

    """
    Создаем вкладку для слияния файлов таблиц
    """
    tab_create_merge_report = ttk.Frame(tab_control)
    tab_control.add(tab_create_merge_report, text='Соединить файлы\n для отчета')

    create_merge_report_frame_description = LabelFrame(tab_create_merge_report)
    create_merge_report_frame_description.pack()

    lbl_hello_create_merge_report = Label(create_merge_report_frame_description,
                                          text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                               'Слияние файлов в общий файл, каждый лист из исходного файла\n'
                                               'будет скопирован на отдельный лист общего файла'
                                          , width=60)
    lbl_hello_create_merge_report.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_merge_report = resource_path('logo.png')
    img_create_merge_report = PhotoImage(file=path_to_img_create_merge_report)
    Label(create_merge_report_frame_description,
          image=img_create_merge_report, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_merge_report = LabelFrame(tab_create_merge_report, text='Подготовка')
    frame_data_merge_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_file_etalon_merge_report = Button(frame_data_merge_report, text='1) Выберите эталонный файл',
                                                 font=('Arial Bold', 14),
                                                 command=select_file_etalon_merge_report)
    btn_choose_file_etalon_merge_report.pack(padx=10, pady=10)

    btn_choose_file_merge_report = Button(frame_data_merge_report, text='2) Выберите папку с исходными файлами',
                                          font=('Arial Bold', 14),
                                          command=select_data_folder_merge_report)
    btn_choose_file_merge_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_merge_report = Button(frame_data_merge_report, text='3) Выберите конечную папку',
                                                font=('Arial Bold', 14),
                                                command=select_end_folder_merge_report)
    btn_choose_end_folder_merge_report.pack(padx=10, pady=10)

    # Создаем кнопку генерации отчетов

    btn_generate_merge_report = Button(tab_create_merge_report, text='4) Соединить таблицы', font=('Arial Bold', 14),
                                       command=processing_merge_report)
    btn_generate_merge_report.pack(padx=10, pady=10)

    """
    Исправление данных для ЕГИССО
    """
    tab_create_fix_data_for_egisso = ttk.Frame(tab_control)
    tab_control.add(tab_create_fix_data_for_egisso, text='Подготовка данных\nдля ЕГИССО')

    create_fix_data_for_egisso_frame_description = LabelFrame(tab_create_fix_data_for_egisso)
    create_fix_data_for_egisso_frame_description.pack()

    lbl_hello_create_fix_data_for_egisso = Label(create_fix_data_for_egisso_frame_description,
                                                 text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                      'Поиск и исправление БАЗОВЫХ ошибок в файлах xlsx для ЕГИССО.\n'
                                                      'Программа удаляет лишние пробелы, лишние символы и т.п.\n'
                                                      'Программа НЕ исправляет неправильные персональные данные\n или идентификаторы ЕГИССО'
                                                 , width=60)
    lbl_hello_create_fix_data_for_egisso.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_fix_data_for_egisso = resource_path('logo.png')
    img_create_fix_data_for_egisso = PhotoImage(file=path_to_img_create_fix_data_for_egisso)
    Label(create_fix_data_for_egisso_frame_description,
          image=img_create_fix_data_for_egisso, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_fix_data_for_egisso = LabelFrame(tab_create_fix_data_for_egisso, text='Подготовка')
    frame_data_fix_data_for_egisso.pack(padx=10, pady=10)

    btn_choose_folder_data_fix_data_for_egisso = Button(frame_data_fix_data_for_egisso,
                                                        text='1) Выберите папку с файлами',
                                                        font=('Arial Bold', 14),
                                                        command=select_folder_data_fix_files_egisso)
    btn_choose_folder_data_fix_data_for_egisso.pack(padx=10, pady=10)

    btn_choose_end_folder_data_fix_data_for_egisso = Button(frame_data_fix_data_for_egisso,
                                                            text='2) Выберите конечную папку',
                                                            font=('Arial Bold', 14),
                                                            command=select_end_folder_data_fix_files_egisso)
    btn_choose_end_folder_data_fix_data_for_egisso.pack(padx=10, pady=10)

    btn_processing_fix_data_for_egisso = Button(frame_data_fix_data_for_egisso,
                                                text='3) Обработать данные',
                                                font=('Arial Bold', 14),
                                                command=processing_fix_files_egisso)
    btn_processing_fix_data_for_egisso.pack(padx=10, pady=10)

    """
    Исправление данных для ГИР ВУ
    """
    tab_create_fix_data_for_girvu = ttk.Frame(tab_control)
    tab_control.add(tab_create_fix_data_for_girvu, text='Подготовка данных\nдля ГИР ВУ')

    create_fix_data_for_girvu_frame_description = LabelFrame(tab_create_fix_data_for_girvu)
    create_fix_data_for_girvu_frame_description.pack()

    lbl_hello_create_fix_data_for_girvu = Label(create_fix_data_for_girvu_frame_description,
                                                text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                     'Поиск и исправление БАЗОВЫХ ошибок в файлах xlsx для ГИР ВУ.\n'
                                                     'Программа удаляет лишние пробелы, лишние символы и т.п.\n'
                                                     'Программа НЕ исправляет неправильные персональные данные\n'
                                                , width=60)
    lbl_hello_create_fix_data_for_girvu.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_fix_data_for_girvu = resource_path('logo.png')
    img_create_fix_data_for_girvu = PhotoImage(file=path_to_img_create_fix_data_for_girvu)
    Label(create_fix_data_for_girvu_frame_description,
          image=img_create_fix_data_for_girvu, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_fix_data_for_girvu = LabelFrame(tab_create_fix_data_for_girvu, text='Подготовка')
    frame_data_fix_data_for_girvu.pack(padx=10, pady=10)

    btn_choose_folder_data_fix_data_for_girvu = Button(frame_data_fix_data_for_girvu,
                                                       text='1) Выберите папку с файлами',
                                                       font=('Arial Bold', 14),
                                                       command=select_folder_data_fix_files_girvu)
    btn_choose_folder_data_fix_data_for_girvu.pack(padx=10, pady=10)

    btn_choose_end_folder_data_fix_data_for_girvu = Button(frame_data_fix_data_for_girvu,
                                                           text='2) Выберите конечную папку',
                                                           font=('Arial Bold', 14),
                                                           command=select_end_folder_data_fix_files_girvu)
    btn_choose_end_folder_data_fix_data_for_girvu.pack(padx=10, pady=10)

    btn_processing_fix_data_for_girvu = Button(frame_data_fix_data_for_girvu,
                                               text='3) Обработать данные',
                                               font=('Arial Bold', 14),
                                               command=processing_fix_files_girvu)
    btn_processing_fix_data_for_girvu.pack(padx=10, pady=10)

    """
        Создаем вкладку для сравнения 2 столбцов
        """
    tab_comparison = Frame(tab_control)
    tab_control.add(tab_comparison, text='Соединение\n2 таблиц')

    comparison_frame_description = LabelFrame(tab_comparison)
    comparison_frame_description.pack()

    lbl_hello_comparison = Label(comparison_frame_description,
                                 text='Соединение 2 таблиц по совпадающим значениям.\nПоиск значений которые есть в обоих таблицах или только в одной.\n'
                                      'ПРИМЕЧАНИЕ\n'
                                      'Заголовок таблиц должен занимать только первую строку!\n'
                                      'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                 width=60)
    lbl_hello_comparison.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_comparison = resource_path('logo.png')
    img_comparison = PhotoImage(file=path_to_img_comparison)
    Label(comparison_frame_description,
          image=img_comparison, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_comparison = LabelFrame(tab_comparison, text='Подготовка')
    frame_data_for_comparison.pack(padx=10, pady=10)

    # Создаем кнопку выбрать файл с параметрами
    btn_columns_params = Button(frame_data_for_comparison, text='1) Выберите файл с параметрами слияния',
                                font=('Arial Bold', 14),
                                command=select_file_params_comparsion)
    btn_columns_params.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_comparison = Button(frame_data_for_comparison, text='2) Выберите первый файл с данными',
                                       font=('Arial Bold', 14),
                                       command=select_first_comparison
                                       )
    btn_data_first_comparison.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name = StringVar()
    # Описание поля
    label_first_sheet_name = Label(frame_data_for_comparison,
                                   text='3) Введите название листа в первом файле')
    label_first_sheet_name.pack(padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_first_sheet_name, width=30)
    first_sheet_name_entry.pack(ipady=5)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_comparison = Button(frame_data_for_comparison, text='4) Выберите второй файл с данными',
                                        font=('Arial Bold', 14),
                                        command=select_second_comparison
                                        )
    btn_data_second_comparison.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name = StringVar()
    # Описание поля
    label_second_sheet_name = Label(frame_data_for_comparison,
                                    text='5) Введите название листа во втором файле')
    label_second_sheet_name.pack(padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry = Entry(frame_data_for_comparison, textvariable=entry_second_sheet_name, width=30)
    second__sheet_name_entry.pack(ipady=5)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_comparison = Button(frame_data_for_comparison, text='6) Выберите конечную папку',
                                       font=('Arial Bold', 14),
                                       command=select_end_folder_comparison
                                       )
    btn_select_end_comparison.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_comparison = Button(tab_comparison, text='7) Произвести слияние\nтаблиц', font=('Arial Bold', 20),
                                    command=processing_comparison
                                    )
    btn_data_do_comparison.pack(padx=10, pady=10)






    """
    Создаем вкладку для проверки истекающих документов
    """
    tab_expired_docs = ttk.Frame(tab_control)
    tab_control.add(tab_expired_docs, text='Заканчивающиеся\n документы')

    expired_docs_frame_description = LabelFrame(tab_expired_docs)
    expired_docs_frame_description.pack()

    lbl_hello_expired_docs = Label(expired_docs_frame_description,
                                   text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                        'Поиск истекающих документов подтверждающих социальные льготы\n'
                                        'Красным выделяются строки если осталось 7 и меньше дней;\n'
                                        'Оранжевым выделяются строки если осталось 14 и меньше дней;\n'
                                        'Желтым выделяются строки если осталось 31 и меньше дней;',
                                   width=60)
    lbl_hello_expired_docs.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_expired_docs = resource_path('logo.png')
    img_expired_docs = PhotoImage(file=path_to_img_expired_docs)
    Label(expired_docs_frame_description,
          image=img_expired_docs, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_expired_docs = LabelFrame(tab_expired_docs, text='Подготовка')
    frame_data_expired_docs.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_prep_file = Button(frame_data_expired_docs, text='1) Выберите файл', font=('Arial Bold', 14),
                                  command=select_file_data_expired)
    btn_choose_prep_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep = Button(frame_data_expired_docs, text='2) Выберите конечную папку',
                                        font=('Arial Bold', 14),
                                        command=select_data_folder_expired)
    btn_choose_end_folder_prep.pack(padx=10, pady=10)

    # Создаем кнопку очистки
    btn_choose_processing_prep = Button(tab_expired_docs, text='3) Выполнить обработку', font=('Arial Bold', 20),
                                        command=processing_check_expired_docs)
    btn_choose_processing_prep.pack(padx=10, pady=10)

    """
         Создаем вкладку для обработки дат рождения
         """

    tab_calculate_date = Frame(tab_control)
    tab_control.add(tab_calculate_date, text='Обработка\nдат рождения')

    calculate_date_frame_description = LabelFrame(tab_calculate_date)
    calculate_date_frame_description.pack()

    lbl_hello_calculate_date = Label(calculate_date_frame_description,
                                     text='Подсчет по категориям,выделение месяца,года\nподсчет текущего возраста,\nподсчет совершеннолетний/несовершеннолетний\n'
                                          'ПРИМЕЧАНИЯ\n'
                                          'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                          'Заголовок таблицы должен занимать только первую строку!\n'
                                          'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                     width=60)
    lbl_hello_calculate_date.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # #
    # #
    # Картинка
    path_to_img_calculate_date = resource_path('logo.png')
    img_calculate_date = PhotoImage(file=path_to_img_calculate_date)
    Label(calculate_date_frame_description,
          image=img_calculate_date, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем фрейм для действий
    calculate_date_frame_action = LabelFrame(tab_calculate_date, text='Подготовка')
    calculate_date_frame_action.pack()

    # Определяем текстовую переменную которая будет хранить дату
    entry_date = StringVar()
    # Описание поля
    label_name_date_field = Label(calculate_date_frame_action,
                                  text='Введите  дату в формате XX.XX.XXXX\n относительно, которой нужно подсчитать текущий возраст\n'
                                       'Например 25.12.2024')
    label_name_date_field.pack(padx=10, pady=10)
    # поле ввода
    date_field = Entry(calculate_date_frame_action, textvariable=entry_date, width=30)
    date_field.pack(ipady=5)

    # Создаем кнопку Выбрать файл с данными
    btn_data_date = Button(calculate_date_frame_action, text='1) Выберите файл с данными', font=('Arial Bold', 14),
                           command=select_file_data_date)
    btn_data_date.pack(padx=10, pady=10)

    btn_choose_end_folder_date = Button(calculate_date_frame_action, text='2) Выберите конечную папку',
                                        font=('Arial Bold', 14),
                                        command=select_end_folder_date
                                        )
    btn_choose_end_folder_date.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_name_column = StringVar()
    # Описание поля
    label_name_column = Label(calculate_date_frame_action,
                              text='3) Введите название колонки с датами рождения,\nкоторые нужно обработать ')
    label_name_column.pack(padx=10, pady=10)
    # поле ввода
    column_entry = Entry(calculate_date_frame_action, textvariable=entry_name_column, width=30)
    column_entry.pack(ipady=5, pady=10)

    btn_calculate_date = Button(tab_calculate_date, text='4) Обработать', font=('Arial Bold', 20),
                                command=calculate_date)
    btn_calculate_date.pack(padx=10, pady=10)

    """
    Извлечение данных из учебных планов для тарификации
    """
    tab_create_extract_data_up = ttk.Frame(tab_control)
    tab_control.add(tab_create_extract_data_up, text='Структурирование\n данных')

    create_extract_data_up_frame_description = LabelFrame(tab_create_extract_data_up)
    create_extract_data_up_frame_description.pack()

    lbl_hello_create_extract_data_up = Label(create_extract_data_up_frame_description,
                                             text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                  'Извлечение данных из таблиц (в том числе из учебных планов)\n'
                                                  'для последующей обработки'
                                             , width=60)
    lbl_hello_create_extract_data_up.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_extract_data_up = resource_path('logo.png')
    img_create_extract_data_up = PhotoImage(file=path_to_img_create_extract_data_up)
    Label(create_extract_data_up_frame_description,
          image=img_create_extract_data_up, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того, чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_extract_data_up = LabelFrame(tab_create_extract_data_up, text='Подготовка')
    frame_data_extract_data_up.pack(padx=10, pady=10)

    # Поля для ввода
    # Определяем текстовую переменную для названия листа
    var_extract_name_sheet = StringVar()
    # Описание поля
    label_extract_name_sheet = Label(frame_data_extract_data_up,
                              text='1) Введите наименование листа на котором располагаются данные')
    label_extract_name_sheet.pack()
    # поле ввода
    entry_extract_name_sheet = Entry(frame_data_extract_data_up, textvariable=var_extract_name_sheet, width=30)
    entry_extract_name_sheet.pack()

    # Определяем текстовую переменную для количества строк заголовка
    var_extract_quantity_header = StringVar()
    # Описание поля
    label_extract_quantity_header = Label(frame_data_extract_data_up,
                              text='2) Введите количество строк, которые занимает заголовок таблицы. Например 4')
    label_extract_quantity_header.pack()
    # поле ввода
    entry_extract_quantity_header = Entry(frame_data_extract_data_up, textvariable=var_extract_quantity_header, width=30)
    entry_extract_quantity_header.pack()

    # Определяем текстовую переменную для порядкового номера колонки которую нужно извлечь
    var_extract_number_main_column = StringVar()
    # Описание поля
    label_extract_number_main_column = Label(frame_data_extract_data_up,
                              text='3) Введите порядковый номер колонки по которой будет происходить обработка.\n'
                                   'Например колонка Предмет является пятой по порядку с начала таблицы, значит нужно ввести число 5')
    label_extract_number_main_column.pack()
    # поле ввода
    entry_extract_number_main_column = Entry(frame_data_extract_data_up, textvariable=var_extract_number_main_column, width=30)
    entry_extract_number_main_column.pack()

    # Определяем текстовую переменную для количества колонок данные которых нужно извлечь
    var_extract_quantity_cols = StringVar()
    # Описание поля
    label_extract_quantity_cols = Label(frame_data_extract_data_up,
                              text='4) Введите количество колонок с данными которые нужно собрать.\n'
                                   'Например если таблица занимает 15 колонок, значит нужно ввести 15')
    label_extract_quantity_cols.pack()
    # поле ввода
    entry_extract_quantity_cols = Entry(frame_data_extract_data_up, textvariable=var_extract_quantity_cols, width=30)
    entry_extract_quantity_cols.pack()

    # Создаем кнопку выбора папки c данными
    btn_choose_data_folder_extract_data_up = Button(frame_data_extract_data_up, text='5) Выберите папку с данными',
                                                   font=('Arial Bold', 14),
                                                   command=select_data_folder_extract_data_up)
    btn_choose_data_folder_extract_data_up.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_extract_data_up = Button(frame_data_extract_data_up, text='6) Выберите конечную папку',
                                                   font=('Arial Bold', 14),
                                                   command=select_end_folder_extract_data_up)
    btn_choose_end_folder_extract_data_up.pack(padx=10, pady=10)
    #
    # Создаем кнопку генерации отчетов

    btn_generate_extract_data_up = Button(tab_create_extract_data_up, text='7) Обработать данные', font=('Arial Bold', 14),
                                          command=processing_extract_data_up)
    btn_generate_extract_data_up.pack(padx=10, pady=10)


    """
    Вкладка для поиска разницы между двумя таблицами
    """
    tab_diffrence = Frame(tab_control)
    tab_control.add(tab_diffrence, text='Разница\n2 таблиц')

    diffrence_frame_description = LabelFrame(tab_diffrence)
    diffrence_frame_description.pack()

    lbl_hello_diffrence = Label(diffrence_frame_description,
                                text='Поиск отличий в двух таблицах\n'
                                     'ВАЖНО Количество строк и колонок в таблицах должно совпадать\n'
                                     'ВАЖНО Названия колонок в таблицах должны совпадать\n'
                                     'ПРИМЕЧАНИЯ\n'
                                     'Заголовок таблицы должен занимать только первую строку!\n'
                                     'Для корректной работы программы уберите из таблицы\n объединенные ячейки',
                                width=60)

    lbl_hello_diffrence.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_diffrence = resource_path('logo.png')
    img_diffrence = PhotoImage(file=path_to_img_diffrence)
    Label(diffrence_frame_description,
          image=img_diffrence, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    frame_data_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_diffrence = Button(frame_data_diffrence, text='1) Выберите первый файл',
                                      font=('Arial Bold', 14),
                                      command=select_first_diffrence
                                      )
    btn_data_first_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_diffrence = Button(frame_data_diffrence, text='2) Выберите второй файл',
                                       font=('Arial Bold', 14),
                                       command=select_second_diffrence
                                       )
    btn_data_second_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_diffrence = Button(frame_data_diffrence, text='3) Выберите конечную папку',
                                      font=('Arial Bold', 14),
                                      command=select_end_folder_diffrence
                                      )
    btn_select_end_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_diffrence = Button(tab_diffrence, text='4) Обработать таблицы', font=('Arial Bold', 20),
                                   command=processing_diffrence
                                   )
    btn_data_do_diffrence.pack(padx=10, pady=10)

    """
      Создаем вкладку создания документов
      """
    tab_create_doc = Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание\nдокументов')

    create_doc_frame_description = LabelFrame(tab_create_doc)
    create_doc_frame_description.pack()

    lbl_hello = Label(create_doc_frame_description,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Генерация документов по шаблону\n'
                           'ПРИМЕЧАНИЯ\n'
                           'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                           'Заголовок таблицы должен занимать только первую строку!\n'
                           'Для корректной работы программы уберите из таблицы\nобъединенные ячейки'
                      , width=60)
    lbl_hello.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(create_doc_frame_description,
          image=img, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем фрейм для действий
    create_doc_frame_action = LabelFrame(tab_create_doc, text='Подготовка')
    create_doc_frame_action.pack()

    # Создаем кнопку Выбрать шаблон
    btn_template_doc = Button(create_doc_frame_action, text='1) Выберите шаблон документа', font=('Arial Bold', 14),
                              command=select_file_template_doc
                              )
    btn_template_doc.pack(padx=10, pady=10)

    btn_data_doc = Button(create_doc_frame_action, text='2) Выберите файл с данными', font=('Arial Bold', 14),
                          command=select_file_data_doc
                          )
    btn_data_doc.pack(padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    # Определяем текстовую переменную
    entry_name_column_data = StringVar()
    # Описание поля
    label_name_column_data = Label(create_doc_frame_action,
                                   text='3) Введите название колонки в таблице\n по которой будут создаваться имена файлов')
    label_name_column_data.pack(padx=10, pady=10)
    # поле ввода
    data_column_entry = Entry(create_doc_frame_action, textvariable=entry_name_column_data, width=30)
    data_column_entry.pack(ipady=5)

    # Поле для ввода названия генериуемых документов
    # Определяем текстовую переменную
    entry_type_file = StringVar()
    # Описание поля
    label_name_column_type_file = Label(create_doc_frame_action, text='4) Введите название создаваемых документов')
    label_name_column_type_file.pack(padx=10, pady=10)
    # поле ввода
    type_file_column_entry = Entry(create_doc_frame_action, textvariable=entry_type_file, width=30)
    type_file_column_entry.pack(ipady=5)

    btn_choose_end_folder_doc = Button(create_doc_frame_action, text='5) Выберите конечную папку',
                                       font=('Arial Bold', 14),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.pack(padx=10, pady=10)

    # Создаем область для того чтобы поместить туда опции
    frame_data_for_options = LabelFrame(tab_create_doc, text='Дополнительные опции')
    frame_data_for_options.pack(padx=10, pady=10)

    # Создаем переменную для хранения переключателя сложного сохранения
    mode_structure_folder_value = StringVar()
    mode_structure_folder_value.set('No')  # по умолчанию сложная структура создаваться не будет
    chbox_mode_structure_folder = Checkbutton(frame_data_for_options,
                                              text='Поставьте галочку, если вам нужно чтобы файлы были сохранены по дополнительным папкам',
                                              variable=mode_structure_folder_value,
                                              offvalue='No',
                                              onvalue='Yes')
    chbox_mode_structure_folder.pack()
    # Создаем поле для ввода
    # Определяем текстовую переменную
    entry_structure_folder_value = StringVar()
    # Описание поля
    label_number_column = Label(frame_data_for_options,
                                text='Введите через запятую не более 3 порядковых номеров колонок по которым будет создаваться структура папок.\n'
                                     'Например: 4,15,8')
    label_number_column.pack()
    # поле ввода
    entry_value_number_column = Entry(frame_data_for_options, textvariable=entry_structure_folder_value, width=30)
    entry_value_number_column.pack(ipady=5)


    # Переключатель краткой версии или полной версии
    mode_full_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_full_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_full = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно чтобы создавались ТОЛЬКО pdf файлы. Работает только в Windows!',
                                       variable=mode_full_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_full.pack()

    # Создаем переменную для хранения результа переключения чекбокса
    mode_combine_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_combine_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_calculate = Checkbutton(frame_data_for_options,
                                       text='Поставьте галочку, если вам нужно чтобы все файлы были объединены в один',
                                       variable=mode_combine_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_calculate.pack()

    # Создаем чекбокс для режима создания pdf
    # Создаем переменную для хранения результа переключения чекбокса
    mode_pdf_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_pdf_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_pdf = Checkbutton(frame_data_for_options,
                                 text='Поставьте галочку, если вам нужно чтобы \n'
                                      'дополнительно создавались pdf версии документов',
                                 variable=mode_pdf_value,
                                 offvalue='No',
                                 onvalue='Yes')
    chbox_mode_pdf.pack()

    # создаем чекбокс для единичного документа

    # Создаем переменную для хранения результа переключения чекбокса
    mode_group_doc_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_group_doc_value.set('No')
    # Создаем чекбокс для выбора режима подсчета
    chbox_mode_group = Checkbutton(frame_data_for_options,
                                   text='Поставьте галочку, если вам нужно создать один документ\nдля конкретного значения (например для определенного ФИО)',
                                   variable=mode_group_doc_value,
                                   offvalue='No',
                                   onvalue='Yes')
    chbox_mode_group.pack(padx=10, pady=10)
    # Создаем поле для ввода значения по которому будет создаваться единичный документ
    # Определяем текстовую переменную
    entry_value_column = StringVar()
    # Описание поля
    label_name_column_group = Label(frame_data_for_options,
                                    text='Введите значение из колонки\nуказанной на шаге 3 для которого нужно создать один документ,\nнапример конкретное ФИО')
    label_name_column_group.pack()
    # поле ввода
    type_file_group_entry = Entry(frame_data_for_options, textvariable=entry_value_column, width=30)
    type_file_group_entry.pack(ipady=5)

    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_doc, text='6) Создать документ(ы)',
                                    font=('Arial Bold', 20),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.pack(padx=10, pady=10)


    """
    Создаем вкладку для предварительной обработки списка
    """
    tab_preparation = ttk.Frame(tab_control)
    tab_control.add(tab_preparation, text='Обработка\nсписка')

    preparation_frame_description = LabelFrame(tab_preparation)
    preparation_frame_description.pack()

    lbl_hello_preparation = Label(preparation_frame_description,
                                  text='Очистка от лишних пробелов и символов; поиск пропущенных значений\n в колонках с персональными данными,'
                                       '(ФИО,паспортные данные,\nтелефон,e-mail,дата рождения,ИНН)\n преобразование СНИЛС в формат ХХХ-ХХХ-ХХХ ХХ.\n'
                                       'Создание списка дубликатов по каждой колонке.\n'
                                       'Поиск со смешаным написанием русских и английских букв.\n'
                                       'ПРИМЕЧАНИЯ\n'
                                       'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                       'Заголовок таблицы должен занимать только первую строку!\n'
                                       'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',
                                  width=60)
    lbl_hello_preparation.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_preparation = resource_path('logo.png')
    img_preparation = PhotoImage(file=path_to_img_preparation)
    Label(preparation_frame_description,
          image=img_preparation, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_prep = LabelFrame(tab_preparation, text='Подготовка')
    frame_data_prep.pack(padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_prep_file = Button(frame_data_prep, text='1) Выберите файл', font=('Arial Bold', 14),
                                  command=select_prep_file)
    btn_choose_prep_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep = Button(frame_data_prep, text='2) Выберите конечную папку', font=('Arial Bold', 14),
                                        command=select_end_folder_prep)
    btn_choose_end_folder_prep.pack(padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса
    mode_dupl_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_dupl_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_dupl = Checkbutton(frame_data_prep,
                                  text='Проверить каждую колонку таблицы на дубликаты',
                                  variable=mode_dupl_value,
                                  offvalue='No',
                                  onvalue='Yes')
    chbox_mode_dupl.pack(padx=10, pady=10)

    # Создаем переменную для хранения результа переключения чекбокса поиска смешения
    mode_mix_alphabets = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_mix_alphabets.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_mix_alphabets = Checkbutton(frame_data_prep,
                                           text='Проверить каждую ячейку таблицы на смешение русских и английских букв',
                                           variable=mode_mix_alphabets,
                                           offvalue='No',
                                           onvalue='Yes')
    chbox_mode_mix_alphabets.pack(padx=10, pady=10)

    # Создаем кнопку очистки
    btn_choose_processing_prep = Button(tab_preparation, text='3) Выполнить обработку', font=('Arial Bold', 20),
                                        command=processing_preparation_file)
    btn_choose_processing_prep.pack(padx=10, pady=10)

    """
    Создание вкладки для разбиения таблицы на несколько штук по значениям в определенной колонке
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_split_tables = ttk.Frame(tab_control)
    tab_control.add(tab_split_tables, text='Разделение\n таблицы')

    split_tables_frame_description = LabelFrame(tab_split_tables)
    split_tables_frame_description.pack()

    lbl_hello_split_tables = Label(split_tables_frame_description,
                                   text='Центр опережающей профессиональной подготовки Республики Бурятия\nРазделение таблицы Excel по листам и файлам'
                                        '\nДля корректной работы программы уберите из таблицы\nобъединенные ячейки\n'
                                        'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                        'Заголовок таблицы должен занимать ОДНУ СТРОКУ\n и в нем не должно быть объединенных ячеек!',
                                   width=60)
    lbl_hello_split_tables.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_split_tables = resource_path('logo.png')
    img_split_tables = PhotoImage(file=path_to_img_split_tables)
    Label(split_tables_frame_description,
          image=img_split_tables, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_split = LabelFrame(tab_split_tables, text='Подготовка')
    frame_data_for_split.pack(padx=10, pady=10)
    # Переключатель:вариант слияния файлов
    # Создаем переключатель
    group_rb_type_split = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_split = LabelFrame(frame_data_for_split, text='1) Выберите вариант разделения')
    frame_rb_type_split.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_type_split, text='А) По листам в одном файле', variable=group_rb_type_split,
                value=0).pack()
    Radiobutton(frame_rb_type_split, text='Б) По отдельным файлам', variable=group_rb_type_split,
                value=1).pack()

    # Создаем кнопку Выбрать файл

    btn_example_split = Button(frame_data_for_split, text='2) Выберите файл с таблицей', font=('Arial Bold', 14),
                               command=select_file_split)
    btn_example_split.pack(padx=10, pady=10)

    # Определяем числовую переменную для порядкового номера
    entry_number_column_split = IntVar()
    # Описание поля
    label_number_column_split = Label(frame_data_for_split,
                                      text='3) Введите порядковый номер колонки начиная с 1\nпо значениям которой нужно разделить таблицу')
    label_number_column_split.pack(padx=10, pady=10)
    # поле ввода имени листа
    entry_number_column_split = Entry(frame_data_for_split, textvariable=entry_number_column_split,
                                      width=30)
    entry_number_column_split.pack(ipady=5)

    btn_choose_end_folder_split = Button(frame_data_for_split, text='4) Выберите конечную папку',
                                         font=('Arial Bold', 14),
                                         command=select_end_folder_split
                                         )
    btn_choose_end_folder_split.pack(padx=10, pady=10)

    # Создаем кнопку слияния

    btn_split_process = Button(tab_split_tables, text='5) Разделить таблицу',
                               font=('Arial Bold', 20),
                               command=processing_split_table)
    btn_split_process.pack(padx=10, pady=10)



    """
    Создаем вкладку для размещения описания программы, руководства пользователя,лицензии.
    """

    tab_about = ttk.Frame(tab_control)
    tab_control.add(tab_about, text='О ПРОГРАММЕ')

    about_frame_description = LabelFrame(tab_about, text='О программе')
    about_frame_description.pack()

    lbl_about = Label(about_frame_description,
                      text="""Деметра - Программа для обработки отчетности ПОО
                           Версия 2.31
                           Язык программирования - Python 3\n
                           Используемая лицензия BSD-2-Clause\n
                           Copyright (c) <2024> <Будаев Олег Тимурович>
                           Адрес сайта программы: https://itdarhan.ru/demetra/demetra.html\n
                           Свидетельство о государственной регистрации № 2024684356
                           
                           Реестровая запись №25751 от 20.12.2024 в реестре 
                           Российского программного обеспечения.

                           Чтобы скопировать ссылку или текст переключитесь на \n
                           английскую раскладку. 
                           """, width=60)

    lbl_about.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    # Картинка
    path_to_img_about = resource_path('logo.png')
    img_about = PhotoImage(file=path_to_img_about)
    Label(about_frame_description,
          image=img_about, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем поле для лицензий библиотек
    guide_frame_description = LabelFrame(tab_about, text='Ссылки для скачивания и обучающие материалы')
    guide_frame_description.pack()

    text_area_url = Text(guide_frame_description, width=84, height=20)
    list_url_path = resource_path('Ссылки.txt')  # путь к файлу лицензии
    with open(list_url_path, 'r', encoding='utf-8') as file:
        text = file.read()
        text_area_url.insert(END, text)
    text_area_url.configure(state='normal')
    text_area_url.pack(side=LEFT)

    scroll = Scrollbar(guide_frame_description, command=text_area_url.yview)
    scroll.pack(side=LEFT, fill=Y)

    text_area_url.config(yscrollcommand=scroll.set)

    text_area_url.configure(state='normal')
    text_area_url.pack(side=LEFT)

    # Кнопка, для демонстрации в отдельном окне списка изменений
    list_changes_path = resource_path('Список изменений.txt')  # путь к файлу лицензии
    button_list_changes = Button(tab_about, text="Список изменений", command=open_list_changes)
    button_list_changes.pack(padx=10, pady=10)

    # Кнопка, для демонстрации в отдельном окне лицензии
    license_path = resource_path('License.txt')  # путь к файлу лицензии
    button_lic = Button(tab_about, text="Лицензия", command=open_license)
    button_lic.pack(padx=10, pady=10)

    # Кнопка, для демонстрации в отдельном окне используемых библиотек
    license_library = resource_path('LibraryLicense.txt')  # путь к файлу с библиотеками
    button_lib = Button(tab_about, text="Дополнительные библиотеки Python", command=open_libraries)
    button_lib.pack(padx=10, pady=10)




    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()
