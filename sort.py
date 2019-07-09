# -*- coding: utf-8 -*-
# Берем из монго, сортируем, результат в Excel

import sys, argparse
from _datetime import datetime, timedelta, date
import os
from mysql.connector import MySQLConnection, Error
from collections import OrderedDict
import openpyxl
from pymongo import MongoClient
from lib import read_config, fine_phone

# подключаемся к серверу
cfg = read_config(filename='anketa.ini', section='Mongo')
conn = MongoClient('mongodb://' + cfg['user'] + ':' + cfg['password'] + '@' + cfg['ip'] + ':' + cfg['port'] + '/'
                   + cfg['db'])
# выбираем базу данных
db = conn.saturn_v

# выбираем коллекцию документов
colls = db.Provider_Alfabank_CreditCards

wb_rez = openpyxl.Workbook(write_only=True)
ws_rez = wb_rez.create_sheet('Даты выгрузки')
ws_rez.append(['Ф', 'И', 'О', 'Телефон', 'Дата и время создания', 'Дата выгрузки'])
for coll in colls.find({},{"passport_lastname": 1, "passport_name" : 1, "passport_middlename" : 1, "personal_phone" : 1,
    "created_date": 1, "history.updated_date" : 1, "history.message" : 1}):
    dates = []
    max_date = datetime(2012, 1, 1, 1, 0, 0)
    updated_dates = coll.get('history', [])
    if len(updated_dates):
        for updated_date in updated_dates:
            if updated_date.get('message', '') == "Заявка выгружена":
                max_date = updated_date['updated_date']
            if updated_date.get('message', '') == "Альфабанк: в обработке.":
                max_date = updated_date['updated_date']
    ws_rez.append([coll['passport_lastname'], coll['passport_name'], coll.get('passport_middlename', ''),
                    fine_phone(coll['personal_phone']), coll['created_date'], max_date])
wb_rez.save('date_upload.xlsx')