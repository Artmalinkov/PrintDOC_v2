'''
Основной функционал приложения по автозамене данных в шаблонах Word
v.2 - пере
'''
import win32api
import os
import datetime
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

# Путь к рабочему файлу Excel
TEST_WORKBOOK_PATH = r'C:\PythonProjects\Print_AS\attachments\excel_tpl.xlsx'

# Наименование таблицы на рабочем листе
TEST_TABLE_NAME = 'Таблица1'

# Путь к рабочему шаблону Word
TEST_DOCTPL_PATH = r"C:\PythonProjects\Print_AS\attachments\word_tpl.docx"

# Столбец по которому определяется необходимость печати документа
FLAG_COLUMN_NAME = 'Печать'

# Загружаем рабочую книгу
wb = load_workbook(TEST_WORKBOOK_PATH)

# Определяем лист
ws = wb.active

# Определяем тяблицу
table = ws.tables[TEST_TABLE_NAME]

# Определяем шаблон Word
doc_tpl = DocxTemplate(TEST_DOCTPL_PATH)

# Необходимость выполнения распечатывания документов
NEED_PRINT = False

# Необходимость сохранения изменённого документа Word в файл
NEED_SAVE = True

# Необходимость замены маркера на сегодняшнюю дату
NEED_CHANGE_NOW_DATE = True


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
        # Преобразование даты
        if type(cell.value) == datetime.datetime:
            cell.value = cell.value.strftime('%d.%m.%Y')

        dict_neces_row[table.column_names[cell.column - 1]] = cell.value

    return dict_neces_row


def save_doc_with_name(changed_doc, context):
    '''
    Функция сохранения измененённого документа
    :param changed_doc: изменённый документ
    :param context: словарь нужной строки
    :return:
    '''
    doc_name = f"C:\PythonProjects\Print_AS\done\{context['Фамилия']}{context['Имя']}.docx"
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


def get_index_neces_column(table, header_name: str) -> int:
    '''
    Функция поиска номера столбца по заданному заголовку
    :param table: объект таблицы на рабочем листе
    :param header_name: заголовок столбца
    :return: номер столбца
    '''
    for header in table.tableColumns:
        if header.name == header_name:
            return header.id


def total_print_doc(worksheet, table, doc_tpl, FLAG_COLUMN_NAME):
    '''
    Сохранение и распечатка каждого документа, отмеченного флагом 'Печать'
    :return:
    '''

    # Определяем, в каком по счёту столбце находится заголовок с именем flag_column_name
    flag_column = index_neces_column(table, FLAG_COLUMN_NAME)

    # Определяем где в столбце flag_column_name стоит '1' - с этой строкой нужно работать.
    for row in worksheet[table.ref]:
        index_neces_row = get_index_neces_row(flag_column, row)
        if index_neces_row != None:
            # Запускаем процедуру замены
            changed_doc, context = one_render(worksheet, table, index_neces_row, doc_tpl)

            # Если необходимо сохранить - сохраняем полученные результаты в файл
            if NEED_SAVE == True:
                save_doc_with_name(changed_doc, context)

            # Если необходимо распечатать - распечатываем полученный файл
            if NEED_PRINT == True:
                print_doc(context)

            if NEED_CHANGE_NOW_DATE == True:
                pass


def get_index_neces_row(index_neces_column, row=row, mark=1):
    '''
    Функция определяет номер нужной строки по ячейке, в которой содержится mark
    :param index_neces_column: номер столбца с нужными ячейками
    :param row: объект исследуемой строки
    :param mark: значение, которое нужно отследить, по умолчанию 1
    :return: index_neces_row
    '''
    if row[index_neces_column].value == mark:
        # Определяем номер строки
        index_neces_row = row[index_neces_column].row - 1
    else:
        index_neces_row = None

    return index_neces_row


def main():
    total_print_doc(ws, table, doc_tpl, FLAG_COLUMN_NAME)


if __name__ == '__main__':
    main()
