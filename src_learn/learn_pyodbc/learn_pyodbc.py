from src.main import *

# Получение объектов соединения и курсора
conn, cursor = connect_to_db()





# Формирование единой строки заголовков без первого столбца
headers_str = ''
for item in headers[1:]:
    headers_str = headers_str + ', ['  + item + ']'
headers_str = headers_str[2:]

headers_str

# Вставка данных в базу данных MS Access
row = ('Сидоров', 'Сидор', 'Сидорович', '07.07.2010')
row = ('Петров')


# Строка SQL-запроса
SQL_str = f'''
        INSERT INTO {table_name} 
        ([Фамилия], [Имя], [Отчество], [Дата_рождения])
        VALUES (?, ?, ?, ?)
        '''

SQL_str = f'''
        INSERT INTO {table_name} 
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
