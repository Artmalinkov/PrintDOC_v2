'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''

import openpyxl
from docxtpl import DocxTemplate



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
import os
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Загрузка тестовой книги и получение таблицы внутри
wb = load_workbook(r'attachments/excel_tpl.xlsx')
ws = wb.active
table = ws.tables['Таблица1']

# Получение списка заголовков. Они находятся в первой строке таблицы
for cell in ws[table.ref][0]:
    print(cell.column)

test_cell = ws.cell(1,1)
test_cell.row

# TODO Получить индекс элемента в списке и продолжить тут
for row in ws[table.ref]:
    print(row)
    for cell in row:
        print (cell.value)
    #print (row)


def get_headers():
    pass