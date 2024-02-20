from src.main import *

# Получение объектов соединения и курсора
conn, cursor = connect_to_db()

# Получение списка заголовков таблицы
table_name = 'Физическое_лицо'
headers = []
for row in cursor.columns(table_name):
    headers.append(row[3])

# Вставка данных в базу данных MS Access
row = ('Сидоров', 'Сидор', 'Сидорович', '07.07.2009')
# Строка SQL-запроса
SQL_str = f'''
        INSERT INTO {table_name} 
        ([Фамилия], [Имя], [Отчество], [Дата_рождения])
        VALUES (?, ?, ?, ?)
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
