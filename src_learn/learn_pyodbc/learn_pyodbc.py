import pyodbc


def get_cursor_connect_to_db(filepath_db):
    '''
    Функция возвращает объект курсора после подключения к базе данных.
    :param filepath_db: Путь к базе данных
    :return: cursor - объект курсора
    '''

    # Создание подключения к Базе данных
    driver_db = 'Microsoft Access Driver (*.mdb, *.accdb)'
    filepath_db = r'C:\PythonProjects\Print_AS\attachments\test_db.accdb'
    user_db = ''
    password_db = ''
    connection_string = f'DRIVER={driver_db};DBQ={filepath_db};UID={user_db};PWD={password_db}'
    conn = pyodbc.connect(connection_string)
    # Cоздание курсора
    cursor = conn.cursor()
    return cursor



# Закрытие подключения
cursor.close()
