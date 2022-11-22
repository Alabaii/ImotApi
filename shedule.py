import datetime
import schedule

from app import create_table


def func():
    print(f'Таблица создана -'+str(datetime.datetime.now()))
    create_table()
    

schedule.every().day.at("00:15").do(func)

while True:
    schedule.run_pending()
