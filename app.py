
from datetime import datetime

from ImotApi import ImotApi
from GsApi import create_calls_report

def yesterday():
    """
    возвращает вчера в формате dd.mm.yyyy

    """
    current_datetime= datetime.now()
    daterange=str(current_datetime.day-1)+'.'+str(current_datetime.month)+'.'+str(current_datetime.year)
    return daterange

def create_table(daterange=yesterday()):
    launch=ImotApi()
    res=launch.get_calls_today(daterange)
    s=launch.get_list_lastnames(res)
    q=launch.get_info_lastname(s,daterange)
    create_calls_report(q)


if __name__=="__main__":
    create_table()