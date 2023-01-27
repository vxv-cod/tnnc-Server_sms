"""Модуль работы с данными на MS SQL 2017 server силами Python"""

import os
import pyodbc
import time
import datetime

from PyQt5 import QtCore, QtWidgets
import os
import sys

from prettytable import PrettyTable
from prettytable import from_db_cursor
from rich import print as rprint

def SMS(Text, title='Выполнено'):
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    QtWidgets.QMessageBox.information(Form, title, Text)

def funCursor():
    adres = "10.28.150.35"
    Database = "TNNC_OAPR_STAT"
    User = "TNNC_OAPR_STAT"
    Password = "RhbgjdsqGfhjkmLkz<L!&$("

    '''Далее создаём строку подключения к нашей базе данных:'''
    connectionString = ("Driver={SQL Server};"
                        f"Server={adres};"
                        f"Database={Database};"
                        f"UID={User};"
                        f"PWD={Password}")

    '''После заполнения строки подключения данными, выполним соединение к нашей базе данных:'''
    connection = pyodbc.connect(connectionString, autocommit=True)

    '''Создадим курсор, с помощью которого, посредством передачи 
    запросов будем оперировать данными в нашей таблице:'''
    dbCursor = connection.cursor()
    return dbCursor

def funInsert(text):
    '''Добавляем значения'''
    dbCursor = funCursor()
    '''Формируем имя пользователя'''
    username = "ROSNEFT\\" + os.getlogin()
    sec = time.localtime(time.time())
    now = f'{sec.tm_mday}-{sec.tm_mon}-{sec.tm_year} {sec.tm_hour}:{sec.tm_min}:{sec.tm_sec}'
    
    '''Добавим данные в нашу таблицу с помощью кода на python:'''
    requestString = f'''INSERT INTO [dbo].StatTable(UserName, ApplicationName, UsingTime) 
                        VALUES  ('{username}', '{text}', '{now}')'''
    dbCursor.execute(requestString)
    SMS(rf'Добавлена запись: "{text}"')

def funUpdate(NameCol, xxx, yyy):
    '''Изменение строк'''
    dbCursor = funCursor()
    '''Обновить таблицу в таблице StatTable в поле(стоблце) ApplicationName на значение 
    замененного ApplicationName с заменой части данных 'УАРМ_' на 'УАРМ'''
    requestString = f'''UPDATE StatTable set {NameCol} = REPLACE ({NameCol}, '{xxx}', '{yyy}')'''
    dbCursor.execute(requestString)
    SMS(rf'Данные {NameCol} заменены с "{xxx}" на "{yyy}" на сервере статистики')

def funDel(user):
    '''Удаление строк'''
    dbCursor = funCursor()
    '''Удаляем записи от пользователя 'ROSNEFT\VVKHOMUTSKIY' '''
    requestString = f'''DELETE FROM StatTable WHERE UserName='{user}';'''
    dbCursor.execute(requestString)
    SMS(rf'Уделение строк "{user}" на сервере статистики')

def visuable():
    mytable = PrettyTable()
    # mytable.field_names = ["ID", "UserName", "ApplicationName", "UsingTime"]
    '''Перебираем записи все записи при условии where'''
    dbCursor = funCursor()
    requestString = f'''SELECT * from StatTable where ApplicationName like(N'%[_]%')'''
    dbCursor.execute(requestString)
    mytable = from_db_cursor(dbCursor)
    # print(from_db_cursor(dbCursor))
    # mytable.align["UserName"] = "l"
    # mytable.align["ApplicationName"] = "l"
    # mytable.border = False
    # print(mytable)
    rprint(mytable)
    # from rich.console import Console
    # from rich import inspect


    # fff = [5, 6, 856, "hdthdhrdrhdh", 000]
    # inspect(fff)


    
    # rows = dbCursor.fetchall()
    # mytable.add_rows(rows)



if __name__ == "__main__":
    # funInsert("Пробный текст УАРМ_000")
    # funUpdate("ApplicationName", 'УАРМ_' , 'УАРМ ')
    # funDel("ROSNEFT\VVKHOMUTSKIY")
    visuable()

    # '''Перебираем записи все записи при условии where'''
    # dbCursor = funCursor()
    # requestString = f'''SELECT * from StatTable where ApplicationName like(N'%[_]%')'''
    # dbCursor.execute(requestString)

    # # rows = dbCursor.fetchall()
    # # mytable.add_rows(rows)

    # print(from_db_cursor(dbCursor))

