# Проработка процедуры вставки данных
# На примере таблицы Физическое лицо
from src.main import *


# Получение объектов соединения и курсора
conn, cursor = connect_to_db()

# Вставка данных в базу данных MS Access
row = ('Сидоров', 'Сидр', 'Петрович', '18.12.2018')


SQL_str = get_SQL_query(cursor, DB_TABLE_NAME)

cursor.execute(SQL_str, row)



print_table_db(cursor, DB_TABLE_NAME)
# Подтвердить изменения
conn.commit()

# Закрыть курсор
cursor.close()
