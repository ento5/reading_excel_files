import os
import openpyxl
import sqlite3

# Создание БД и подключение к ней
conn_db = sqlite3.connect("list_db")

# Создание курсора - это специальный объект который делает запросы и получает их результаты
cursor = conn_db.cursor()

# Создание таблицы в БД
cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fio TEXT,
    birthday TEXT,
    registration TEXT,
    inn INTEGER,
    issued TEXT
)
""")

# Указание пути файла где он находиться для дальнейшего считывания
path = "C:/Users/ant74/Desktop/IT/SQL/Данные для БД/users_excel.xlsx"

# Чтение и загрузка книги
read_file = openpyxl.load_workbook(path)
sheet = read_file.active

# Чтение данных из ячеек по строкам
for row in sheet.iter_rows(values_only=True):
# Запись данных в БД которые получили выше и записали в row
    cursor.execute('''
        INSERT INTO users (fio, birthday, registration, inn, issued) VALUES
(?, ?, ?, ?, ?)
    ''', row)
conn_db.commit()
conn_db.close()

print("Данные успешно сохранены в базу данных.")
