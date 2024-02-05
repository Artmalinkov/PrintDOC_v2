'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

TEST_WORKBOOK_PATH = r'attachments/excel_tpl.xlsx'
TEST_TABLE_NAME = 'Таблица1'

# Загружаем рабочую книгу
wb = load_workbook(TEST_WORKBOOK_PATH)

# Определяем лист
ws = wb.active

# Определяем тяблицу
table = ws.tables[TEST_TABLE_NAME]


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


def get_headers(path_to_book: str, table_name: str):
    '''
    Формирование словаря где ключ - название поля таблицы, а значение - номер столбца
    :param path_to_book: путь к рабочей книге excel
    :param table_name: наименование таблицы
    :return: dict_headers - словарь заголовков таблицы
    '''
    # Загрузка тестовой книги и получение таблицы внутри
    wb = load_workbook(path_to_book)
    ws = wb.active
    table = ws.tables[table_name]

    # Получение списка заголовков. Они находятся в первой строке таблицы
    dict_headers = {}
    for cell in ws[table.ref][0]:
        dict_headers[cell.value] = cell.column

    # TODO Проработать вынесение функционала загрузки рабочей униги отдельно
    # TODO Прорпботать формирование этого словаря через атрибут column_names объекта table
    # TODO Прорпботать формирование этого словаря через атрибут tableColumns объекта table

    return dict_headers


def get_neces_row(worksheet, table, index_neces_row: int):
    '''
    Формирование словаря, где ключ - название поля таблицы, а значение - ячейка в строке с индексом 'index_neces_row'
    :param worksheet: активный лист
    :param table: Таблица Excel
    :param index_neces_row: индекс нужной строки
    :return: словарь нужной строки
    '''
    # определяем пустрой словарь
    dict_neces_row = {}
    # table.ref - диапазон умной таблицы
    # column_names - список заголовков умной таблицы
    # cell.column - индекс столбца ячейки
    # cell.column - индекс столбца ячейки
    # cell.value - значение ячейки
    for cell in worksheet[table.ref][index_neces_row]:
        dict_neces_row[table.column_names[cell.column - 1]] = cell.value

    return dict_neces_row
