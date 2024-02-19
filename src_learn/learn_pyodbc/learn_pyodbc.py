import pyodbc


# Срздание подключения к Базе данных
driver_db = 'Microsoft Access Driver (*.mdb, *.accdb)'
filepath_db = r'C:\PythonProjects\Print_AS\src_learn\learn_pyodbc\test_db.accdb'
user_db = ''
password_db = ''
connection_string = f'DRIVER={{{driver_db}}};DBQ={filepath_db};UID={user_db};PWD={password_db}'
conn = pyodbc.connect(connection_string)

# Cоздание курсора
cursor = conn.cursor()


