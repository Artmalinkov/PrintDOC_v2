# Проработка процедуры вставки данных
# На примере таблицы Физическое лицо
from src.main import *


# Получение объектов соединения и курсора
conn, cursor = connect_to_db()








symb

# Вставка данных в базу данных MS Access
row = ('Сидоров', 'Сидор', 'Сидорович', '07.07.2010')
row = ('Петров', 'Петр', 'Петрович', None)


def get_SQL_query(cursor, DB_TABLE_NAME):
    '''
    Функция возвращает переработанную строку SQL-запроса на основании объекта курсора и наименовании таблицы
    :param cursor: объект курсора базы данных
    :param DB_TABLE_NAME: наименование таблицы в базе данных
    :return: SQL_str - строка запроса
    '''
    # Получение списка заголовка в таблице
    headers = get_headers(cursor, DB_TABLE_NAME)

    # Определение количества вопросительных знаков
    symb = ('?,' * (len(headers) - 1))[:-1]

    # Формирование строки заголовков в нужном формате
    headers_str = get_headers_str(cursor, DB_TABLE_NAME)

    # Формирование итоговой строки SQL-запроса
    SQL_str = f'INSERT INTO {DB_TABLE_NAME} ({headers_str}) VALUES ({symb})'

    return SQL_str



sql = get_SQL_query(cursor, DB_TABLE_NAME)

cursor.execute(SQL_str, row)

# Выборка для проверки результата
cursor.execute('SELECT * FROM Физическое_лицо')
# Вывести на экран результаты
for row in cursor.fetchall():
    print(row)

# Подтвердить изменения
conn.commit()

# Закрыть курсор
cursor.close()
