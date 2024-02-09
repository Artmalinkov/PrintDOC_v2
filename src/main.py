'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
import os
import datetime
import win32api
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

# Путь к рабочему файлу Excel
TEST_WORKBOOK_PATH = r'C:\PythonProjects\Print_AS\attachments\excel_tpl.xlsx'

# Наименование таблицы на рабочем листе
TABLE_NAME = 'Таблица1'

# Столбец по которому определяется необходимость печати документа
PRINT_COLUMN_NAME = 'Печать'

# Маркер по отслеживанию необходимости печати
MARK = 1

# Путь к рабочему шаблону Word
TEST_DOCTPL_PATH = r"C:\PythonProjects\Print_AS\attachments\word_tpl.docx"

# Путь к папке, в которой будут лежать результаты
RESULT_DOC_DIR = r"C:\PythonProjects\Print_AS\done"

# Загружаем рабочую книгу
wb = load_workbook(TEST_WORKBOOK_PATH)

# Определяем лист
ws = wb.active

# Определяем тяблицу
table = ws.tables[TABLE_NAME]
#
# Определяем шаблон Word
doc_tpl = DocxTemplate(TEST_DOCTPL_PATH)

# Необходимость выполнения распечатывания документов
NEED_PRINT = True

# Необходимость сохранения изменённого документа Word в файл
NEED_SAVE = True

# Необходимость замены маркера на сегодняшнюю дату
NEED_CHANGE_NOW_DATE = True


def get_dict_row(table, row):
    '''
    Функция формирования словаря из значений в строке
    :param table: объект таблицы на рабочем листе
    :param row: объект строки в таблице
    :return: dict_row - сформированный словарь: ключ - заголовок столбца, значение - ячейка
    '''
    dict_row = {}
    for cell in row:
        # Преобразование даты
        if type(cell.value) == datetime.datetime:
            cell.value = cell.value.strftime('%d.%m.%Y')
        # Наполнение словаря
        dict_row[table.column_names[cell.column - 1]] = cell.value
    return dict_row


def iteration_row(MARK, ws, table):
    '''
    Функция просмотра строк на рабочем листе в таблице.
    :param MARK: маркер по котороку определяется необходимость печати
    :param ws: рабочий лист
    :param table: рабочая таблица
    :return:
    '''

    # Запускаем процедуру просмотра строк в таблице.
    for row in ws[table.ref]:

        # Для каждой строки формируем словарь
        dict_row = get_dict_row(table, row)

        # Если в нужном поле стоит единичка...
        if dict_row[PRINT_COLUMN_NAME] == MARK:

            # Запускаем процедуру замены
            doc_tpl.render(dict_row)

            # Если нужно - сохраняем, потом если нужно - распечатываем.
            if NEED_SAVE == True:
                save_doc_with_name(doc_tpl, dict_row)

                # Если нужно - распечатываем
                if NEED_PRINT == True:
                    print_doc(dict_row)


def save_doc_with_name(changed_doc, dict_row):
    '''
    Функция сохранения измененённого документа
    :param changed_doc: изменённый документ
    :param context: словарь нужной строки
    :return:
    '''
    doc_name = RESULT_DOC_DIR + '\\' + f"{dict_row['Фамилия']}{dict_row['Имя']}.docx"
    changed_doc.save(doc_name)


def print_doc(dict_row):
    '''
    Функция распечатки документа
    :param dict_row: словарь нужной строки
    :return:
    '''

    doc_name = f"done/{dict_row['Фамилия']}{dict_row['Имя']}.docx"

    full_filepath = os.path.abspath(doc_name)
    win32api.ShellExecute(0, 'print', full_filepath, None, '.', 0)


def print_doc(dict_row):
    '''
    Функция распечатки документа
    :param dict_row: словарь нужной строки
    :return:
    '''
    full_filepath = RESULT_DOC_DIR + '\\' + f"{dict_row['Фамилия']}{dict_row['Имя']}.docx"
    win32api.ShellExecute(0, 'print', full_filepath, None, '.', 0)


def main():
    iteration_row(MARK, ws, table)
    pass


if __name__ == '__main__':
    main()
