"""Модуль отправки данный на MS SQL 2017 server силами Python"""

import os
import pyodbc
import time

def dataSQL():
    try:
        '''Читаем файл и разбивает строку на подстроки в зависимости от разделителя'''
        f = open("tnnc_oapr_stat_text.ini", 'r', encoding='utf-8')
        datalist = f.read().split("\n")
        datalist = datalist[:3]
        f.close()
    except:
        print("Файл не найден......")
        datalist = ["10.28.150.35", "Пробный текст . . .", "Отправлять текст адреса папки по индексу? - НЕТ"]

    '''Собираем адресс сервера и текст отправки на него'''
    YESS = "Отправлять текст адреса папки по индексу? - ДА"
    '''Отправляем текст с названием папки'''
    if YESS in datalist[2]:
        nomer = int(datalist[2][-2:])
        pathUaarm = os.getcwd()
        text = pathUaarm.split('\\')[nomer]
        datalist[1] = text
    # print('datalist = ', datalist)
    return datalist

def Sql():
    '''Формируем текущую дату'''
    sec = time.localtime(time.time())
    now = f'{sec.tm_mday}-{sec.tm_mon}-{sec.tm_year} {sec.tm_hour}:{sec.tm_min}:{sec.tm_sec}'

    '''Формируем имя пользователя'''
    username = "ROSNEFT\\" + os.getlogin()
    
    datalist = dataSQL()
    adres = datalist[0]
    text = datalist[1]

    '''Далее создаём строку подключения к нашей базе данных:'''
    connectionString = ("Driver={SQL Server};"
                        # "Server=10.28.150.35;"
                        f"Server={adres};"
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
    Sql()