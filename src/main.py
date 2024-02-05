'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
import os
from docxtpl import DocxTemplate
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

TEST_WORKBOOK_PATH = r'attachments/excel_tpl.xlsx'
TEST_TABLE_NAME = 'Таблица1'


# Открытие и редактор шаблона Word
def replace_Word_doc():
    '''
        Замена полей в шаблоне Word
        :return: None
        '''
    # TODO проработать передачу пути
    # TODO проработать передачу словаря значений
    doc = DocxTemplate(r"attachments/word_tpl.docx")
    context = {
        'переменная': 'Название компании'
    }
    doc.render(context)
    doc.save(r"done/generated_docx.docx")


# формирование словаря где ключ - название поля таблицы, а значение - номер столбца
def get_headers(path_to_book: str,  table_name: str, ):
    # Загрузка тестовой книги и получение таблицы внутри
    wb = load_workbook(path_to_book)
    ws = wb.active
    table = ws.tables[table_name]

    # Получение списка заголовков. Они находятся в первой строке таблицы
    dict_headers = {}
    for cell in ws[table.ref][0]:
        dict_headers[cell.value] = cell.column

    return dict_headers




test_dict_headers = get_headers(TEST_WORKBOOK_PATH, TEST_TABLE_NAME)
test_dict_headers