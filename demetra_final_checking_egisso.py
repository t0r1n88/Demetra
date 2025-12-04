"""
Скрипт для финальной выверки данных ЕГИССО
"""
import numpy as np

from demetra_support_functions import write_df_to_excel_cheking_egisso,del_sheet,convert_to_date_egisso_cheking,create_doc_convert_date_egisso_cheking,convert_to_date_start_finish_egisso_cheking,write_df_error_egisso_to_excel # вспомогательные функции
import os
import pandas as pd
from tkinter import messagebox
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import xlsxwriter
import time
from datetime import datetime
import re
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import logging
logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',)


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


class NotRecColsLMSZ(Exception):
    """
    Обработка случаев когда нет обязательных колонок в файле
    """
    pass











def final_checking_files_egisso(data_folder:str, end_folder:str):
    """
    Функция для выверки данных ЕГИССО
    :param data_folder: папка с данными
    :param end_folder: конечная папка
    """








if __name__ == '__main__':
    main_data_folder = 'c:/Users/1/PycharmProjects/Demetra/data/ЕГИССО'
    main_end_folder = 'c:/Users/1/PycharmProjects/Demetra/data/СБОР результат'

    start_time = time.time()
    final_checking_files_egisso(main_data_folder, main_end_folder)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Время выполнения: {elapsed_time:.6f} сек.")


    print('Lindy Booth')
