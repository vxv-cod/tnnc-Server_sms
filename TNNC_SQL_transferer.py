'''
Автонмный модуль для отправки данный на MS SQL 2017 server
особенности: ищем название УАРМ из UARMWindow.xml рядом лежащим
'''
import os
import time
import pyodbc

def Sql(text):
    print("-----------------")
    '''Формируем текущую дату'''
    sec = time.localtime(time.time())
    now = f'{sec.tm_mday}-{sec.tm_mon}-{sec.tm_year} {sec.tm_hour}:{sec.tm_min}:{sec.tm_sec}'

    '''Формируем имя пользователя'''
    username = "ROSNEFT\\" + os.getlogin()
    
    '''Далее создаём строку подключения к нашей базе данных:'''
    connectionString = ("Driver={SQL Server};"
                        "Server=10.28.150.35;"
                        "Database=TNNC_OAPR_STAT;"
                        "UID=TNNC_OAPR_STAT;"
                        "PWD=RhbgjdsqGfhjkmLkz<L!&$(")

    '''После заполнения строки подключения данными, выполним соединение к нашей базе данных:'''
    connection = pyodbc.connect(connectionString, autocommit=True)

    '''Создадим курсор, с помощью которого, посредством передачи 
    запросов будем оперировать данными в нашей таблице:'''
    dbCursor = connection.cursor()

    '''Добавим данные в нашу таблицу с помощью кода на python:'''
    requestString = f'''INSERT INTO [dbo].StatTable(UserName, ApplicationName, UsingTime) 
                        VALUES  ('{username}', '{text}', '{now}')'''
    dbCursor.execute(requestString)

    '''Сохранение данный в базе'''
    connection.commit()

if __name__ == "__main__":
    '''Берем текст из файла'''
    f = open("UARMWindow.xml", encoding='utf-8')
    xxx = f.read().split("\n")

    if "<title>" in xxx[-2]:
        text = xxx[-2].split("<title>")[1].split("</title>")[0]
    else:
        text = xxx[-3].split("<title>")[1]

    f.close()
    
    print(f"text = {text}")
    Sql(text)

