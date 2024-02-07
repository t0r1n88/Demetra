"""
Графический интерфейс для программы
"""
from create_local_report import create_local_report # создание отчета по выбранным пользователем параметрам
from create_social_passport import create_social_report # создание отчета по социальному состоянию
from preparation_list import prepare_list # подготовка персональных данных
from split_table import split_table # разделение таблицы
import pandas as pd
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

"""
Системные функции
"""
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
        width = int(screen_width * 0.41)
    elif screen_width >= 1600:
        width = int(screen_width * 0.5)
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
def select_file_params_local_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_params_local_report
    # Получаем путь к файлу
    name_file_params_local_report = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_file_data_local_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_data_local_report
    # Получаем путь к файлу
    name_file_data_local_report = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


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
        create_local_report(name_file_data_local_report,path_to_end_folder_local_report,name_file_params_local_report)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента','Выберите файл с параметрами,файл с данными, конечную папку')


"""
Создание социального отчета по контингенту БРИТ
"""

def select_file_data_social_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться локальные отчеты соцпедагога
    :return: Путь к файлу с данными
    """
    global name_file_data_social_report
    # Получаем путь к файлу
    name_file_data_social_report = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


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
        create_social_report(name_file_data_social_report,path_to_end_folder_social_report)
    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента','Выберите файл с параметрами,файл с данными, конечную папку')


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
        number_column = entry_number_column_split.get() #  получаем порядковый номер колонки
        number_column = int(number_column) # конвертируем в инт

        checkbox_split = group_rb_type_split.get() # получаем значения переключиталея

        # находим разницу
        split_table(file_data_split,number_column,checkbox_split,path_to_end_folder_split)
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
        prepare_list(glob_prep_file,glob_path_to_end_folder_prep,checkbox_dupl)

    except NameError:
        messagebox.showerror('Деметра Отчеты социальный паспорт студента',
                             f'Выберите файл с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')



if __name__ == '__main__':
    window = Tk()
    window.title('Деметра Отчеты социальный паспорт студента ver 1.0')
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

    """
       Создаем вкладку для создания социального паспорта БРИТ
       """
    tab_create_social_report = ttk.Frame(tab_control)
    tab_control.add(tab_create_social_report, text='Социальный паспорт')

    create_social_report_frame_description = LabelFrame(tab_create_social_report)
    create_social_report_frame_description.pack()

    lbl_hello_create_social_report = Label(create_social_report_frame_description,
                                           text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                'Создание отчета по социальному статусу контингента'
                                           , width=60)
    lbl_hello_create_social_report.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_create_social_report = resource_path('logo.png')
    img_create_social_report = PhotoImage(file=path_to_img_create_social_report)
    Label(create_social_report_frame_description,
          image=img_create_social_report, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_social_report = LabelFrame(tab_create_social_report, text='Подготовка')
    frame_data_social_report.pack(padx=10, pady=10)

    btn_choose_file_social_report = Button(frame_data_social_report, text='1) Выберите файл', font=('Arial Bold', 14),
                                           command=select_file_data_social_report)
    btn_choose_file_social_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_social_report = Button(frame_data_social_report, text='2) Выберите конечную папку',
                                                 font=('Arial Bold', 14),
                                                 command=select_end_folder_social_report)
    btn_choose_end_folder_social_report.pack(padx=10, pady=10)

    # Создаем кнопку генерации отчетов

    btn_generate_social_report = Button(tab_create_social_report, text='3) Создать отчеты', font=('Arial Bold', 14),
                                        command=processing_social_report)
    btn_generate_social_report.pack(padx=10, pady=10)

    """
    Создаем вкладку для создания управляемых  по социальному контингенту БРИТ
    """
    tab_create_local_report= ttk.Frame(tab_control)
    tab_control.add(tab_create_local_report, text='Настраиваемый отчет')

    create_local_report_frame_description = LabelFrame(tab_create_local_report)
    create_local_report_frame_description.pack()

    lbl_hello_create_local_report = Label(create_local_report_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                       'Создание настраиваемых отчетов для соцпедагога'
                                       ,width=60)
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

    # Создаем кнопку выбора файла с данными
    btn_choose_file_params_local_report= Button(frame_data_local_report, text='1) Выберите файл c параметрами', font=('Arial Bold', 14),
                                       command=select_file_params_local_report)
    btn_choose_file_params_local_report.pack(padx=10, pady=10)


    btn_choose_file_local_report= Button(frame_data_local_report, text='2) Выберите файл', font=('Arial Bold', 14),
                                       command=select_file_data_local_report)
    btn_choose_file_local_report.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_local_report= Button(frame_data_local_report, text='3) Выберите конечную папку', font=('Arial Bold', 14),
                                       command=select_end_folder_local_report)
    btn_choose_end_folder_local_report.pack(padx=10, pady=10)

    # Создаем кнопку генерации отчетов

    btn_generate_local_report = Button(tab_create_local_report,text='4) Создать отчеты', font=('Arial Bold', 14),command=processing_local_report)
    btn_generate_local_report.pack(padx=10, pady=10)


    """
    Создаем вкладку для предварительной обработки списка
    """
    tab_preparation= ttk.Frame(tab_control)
    tab_control.add(tab_preparation, text='Подготовка списка')

    preparation_frame_description = LabelFrame(tab_preparation)
    preparation_frame_description.pack()

    lbl_hello_preparation = Label(preparation_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                       'Очистка от лишних пробелов и символов; поиск пропущенных значений\n в колонках с персональными данными,'
                                       '(ФИО,паспортные данные,\nтелефон,e-mail,дата рождения,ИНН)\n преобразование СНИЛС в формат ХХХ-ХХХ-ХХХ ХХ.\n'
                                       'Создание списка дубликатов по каждой колонке\n'
                                       'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                       'Для корректной работы программы уберите из таблицы\nобъединенные ячейки',width=60)
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
    btn_choose_prep_file= Button(frame_data_prep, text='1) Выберите файл', font=('Arial Bold', 14),
                                       command=select_prep_file)
    btn_choose_prep_file.pack(padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep= Button(frame_data_prep, text='2) Выберите конечную папку', font=('Arial Bold', 14),
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


    # Создаем кнопку очистки
    btn_choose_processing_prep= Button(tab_preparation, text='3) Выполнить обработку', font=('Arial Bold', 20),
                                       command=processing_preparation_file)
    btn_choose_processing_prep.pack(padx=10, pady=10)

    """
    Создание вкладки для разбиения таблицы на несколько штук по значениям в определенной колонке
    """
    # Создаем вкладку для подсчета данных по категориям
    tab_split_tables = ttk.Frame(tab_control)
    tab_control.add(tab_split_tables, text='Разделение таблицы')

    split_tables_frame_description = LabelFrame(tab_split_tables)
    split_tables_frame_description.pack()

    lbl_hello_split_tables = Label(split_tables_frame_description,
                                   text='Центр опережающей профессиональной подготовки Республики Бурятия\nРазделение таблицы Excel по листам и файлам'
                                       '\nДля корректной работы программы уберите из таблицы\nобъединенные ячейки\n'
                                       'Данные обрабатываются С ПЕРВОГО ЛИСТА В ФАЙЛЕ !!!\n'
                                       'Заголовок таблицы должен занимать ОДНУ СТРОКУ\n и в нем не должно быть объединенных ячеек!',width=60)
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