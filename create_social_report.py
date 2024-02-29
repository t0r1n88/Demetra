"""
Скрипт по созданию отчета по социальному положению БРИТ
"""

from support_functions import *
import pandas as pd
import openpyxl
from copy import copy
import time
from collections import Counter
import re
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None



def create_social_report(etalon_file:str, folder_update_file:str, result_folder:str):
    """
    Функция для оздания отчета по социальному положению ГБПОУ БРИТ
    :param etalon_file: файл с которым будут сравниваться данные
    :param folder_update_file: папка в которой лежат файлы для обновления
    :param result_folder: папка в которой будет находится итоговый файл
    """







if __name__=='__main__':
    main_etalon_file = 'data/Таблица для заполнения социального паспорта студентов.xlsx'
    main_folder_update = 'data/27.02'
    main_folder_result = 'data/Результат'
    merge_table(main_etalon_file,main_folder_update,main_folder_result)

    print('Lindy Booth')