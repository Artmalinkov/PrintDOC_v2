'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
import win32api
import os

# Путь к рабочему файлу Excel
TEST_WORKBOOK_PATH = r'attachments/excel_tpl.xlsx'

# Наименование таблицы на рабочем листе
TEST_TABLE_NAME = 'Таблица1'

# Путь к рабочему шаблону Word
TEST_DOCTPL_PATH = r"attachments/word_tpl.docx"

# Путь по которуму нужно сохранять преобразованный файл Word
# TODO убрать, поскольку именование происходит по Фамилии и имени
DOC_PATH_SAVE = r"done/generated_docx.docx"

# Наименование столбца, по которому отслеживается печать
FLAG_COLUMN_NAME = 'Печать'

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


def save_doc_with_name(changed_doc, context):
    '''
    Функция сохранения измененённого документа
    :param changed_doc: изменённый документ
    :param context: словарь нужной строки
    :return:
    '''
    doc_name = f"done/{context['Фамилия']}{context['Имя']}.docx"
    changed_doc.save(doc_name)


def print_doc(context):
    '''
    Функция распечатки документа
    :param context: словарь нужной строки
    :return:
    '''
    doc_name = f"done/{context['Фамилия']}{context['Имя']}.docx"
    full_filepath = os.path.abspath(doc_name)
    win32api.ShellExecute(0, 'print', full_filepath, None, '.', 0)


def get_column_id(flag_column_name: str) -> int:
    '''
    Функция поиска номера столбца по заданному заголовку
    :param flag_column_name: заголовок столбца
    :return: номер столбца
    '''
    for header in table.tableColumns:
        if header.name == flag_column_name:
            flag_column_id = header.id
    return flag_column_id


def total_print_doc():
    '''
    Сохранение и распечатка каждого документа, отмеченного флагом 'Печать'
    :return:
    '''
    # Определяем, в каком по счёту столбце находится заголовок с именем flag_column_name
    flag_column = get_column_id('Печать')

    # Определяем где в столбце flag_column_name стоит '1' - с этой строкой нужно работать.
    for row in ws[table.ref]:
        if row[flag_column - 1].value == 1:
            # Определяем номер строки
            index_neces_row = row[flag_column - 1].row - 1

            # Запускаем процедуру замены
            changed_doc, context = one_render(ws, table, index_neces_row, doc_tpl)

            # Сохраняем полученные результаты в файл
            save_doc_with_name(changed_doc, context)

            # Распечатываем полученный файл
            print_doc(context)


def main():
    pass


if __name__ == '__main__':
    main()

