from src.main import *

# Получение объектов соединения и курсора
conn, cursor = connect_to_db()



headers_str = get_headers_str(cursor, DB_TABLE_NAME)




# Вставка данных в базу данных MS Access
row = ('Сидоров', 'Сидор', 'Сидорович', '07.07.2010')
row = ('Петров', 'Петр', 'Петрович', None)


# Строка SQL-запроса
SQL_str = f'''
        INSERT INTO {DB_TABLE_NAME} 
        ({headers_str})
        VALUES (?,?,?,?)
        '''

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
