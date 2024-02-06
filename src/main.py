'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

TEST_WORKBOOK_PATH = r'attachments/excel_tpl.xlsx'
TEST_TABLE_NAME = 'Таблица1'
TEST_DOCTPL_PATH = r"attachments/word_tpl.docx"
DOC_PATH_SAVE = r"done/generated_docx.docx"

# Загружаем рабочую книгу
wb = load_workbook(TEST_WORKBOOK_PATH)

# Определяем лист
ws = wb.active

# Определяем тяблицу
table = ws.tables[TEST_TABLE_NAME]

# Определяем шаблон Word
doc_tpl = DocxTemplate(TEST_DOCTPL_PATH)


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


def one_render(ws, table, index_neces_row, doc_tpl):
    '''
    Функция единичной замены в шаблоне Word и сохранения в новый готовый документ
    :param ws: рабочий лист Excel
    :param table: таблица на листе
    :param index_neces_row: номер нужной строки для замены
    :param doc_tpl: шаблон документа Word
    :return: doc_tpl: изменённый документ Word,
             contex: словарь нужной строки
    '''
    context = get_neces_row(ws, table, index_neces_row)
    doc_tpl.render(context)
    # TODO Проработать преобразование формат дат.
    return doc_tpl, context


def save_doc_with_name(ws, table, index_neces_row):
    '''
    Функция сохранения измененённого документа
    :param ws: рабочий лист
    :param table: страница на листе
    :param index_neces_row: номер строки в таблице
    :return:
    '''
    changed_doc, context = one_render(ws, table, index_neces_row, doc_tpl)
    doc_name = f"done/{context['Фамилия']}{context['Имя']}.docx"
    changed_doc.save(doc_name)

# Поскольку заголовки можно получать напрямую из атрибутов таблицы, словарь заголовков пока не нужен.
# def get_headers(path_to_book: str, table_name: str):
#     '''
#     Формирование словаря где ключ - название поля таблицы, а значение - номер столбца
#     :param path_to_book: путь к рабочей книге excel
#     :param table_name: наименование таблицы
#     :return: dict_headers - словарь заголовков таблицы
#     '''
#     # Загрузка тестовой книги и получение таблицы внутри
#     wb = load_workbook(path_to_book)
#     ws = wb.active
#     table = ws.tables[table_name]
#
#     # Получение списка заголовков. Они находятся в первой строке таблицы
#     dict_headers = {}
#     for cell in ws[table.ref][0]:
#         dict_headers[cell.value] = cell.column
#
#     # TODO Проработать вынесение функционала загрузки рабочей униги отдельно
#     # TODO Прорпботать формирование этого словаря через атрибут column_names объекта table
#     # TODO Прорпботать формирование этого словаря через атрибут tableColumns объекта table
#
#     return dict_headers
