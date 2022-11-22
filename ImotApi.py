import requests
import json
from datetime import datetime
from pprint import pprint


class Req( object ):
    """
    обертка аутентификации

    """

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    acc_tkn = ''

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    def __init__( self ):

        super( Req, self ).__init__()
        self.req_ini()

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    def req_ini( self ):

        self.acc_tkn = 'e48f31ac-dbaf-4ae6-b691-2d3a842ac06a'

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    def post( self, _url_, _dta_ ):

        tkn = 'e48f31ac-dbaf-4ae6-b691-2d3a842ac06a'

        hdr = { 'X-Auth-Token':tkn }
        url = _url_
        dta = _dta_

        try    : res = requests.post( url, headers=hdr, data=dta )
        except : res = requests.post( url, headers=hdr, json=dta )

        return res

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    def put( self, _url_, _dta_ ):

        tkn = 'e48f31ac-dbaf-4ae6-b691-2d3a842ac06a'

        hdr = { 'X-Auth-Token':tkn }
        url = _url_
        dta = _dta_

        try    : res = requests.put( url, headers=hdr, data=dta )
        except : res = requests.put( url, headers=hdr, json=dta )

        return res

# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~

    def get( self, _url_ ):

        tkn = 'e48f31ac-dbaf-4ae6-b691-2d3a842ac06a'

        hdr = { 'X-Auth-Token':tkn }
        url = _url_

        res = requests.get( url, headers=hdr )

        return res



class ImotApi:
    _MAIN_URL = "https://test.imot.io/api/"

    def __init__(self):
        self.req = Req()   # -- main request object
    
    def get_calls_today(self, daterange):
        """
        Возвращает список строк всех id звонков за вчерашний день

        """
        result = self.req.get (f"{self._MAIN_URL}calls?daterange="+daterange)
        calls_today_json = result.text
        list_calls = json.loads(calls_today_json)
        total_calls = len(list_calls)
        i=0
        id_calls=[]
        while i < total_calls:
            id_calls.append(list_calls[i]['id'])
            i += 1
        return id_calls

    def get_list_lastnames(self, id):
        """

        Возвращает словарь формата 
        {
            'ФИО' : ['список id звонков'],
            '':['','','','']
        }
        
        """
        
        lastnames = []
        for item in id:
            res = Req().get(f"{self._MAIN_URL}call/"+item+"/tags" )
            tags=json.loads(res.text)
            employee=next((x for x in tags if x['name'] == 'Сотрудник'), None)
            lastnames.append(employee['value'])
        id_and_lastnames=dict(zip(id, lastnames))
        lastnames_and_id={}
        for k, v in id_and_lastnames.items():
            lastnames_and_id[v] = lastnames_and_id.get(v, []) + [k]
        return lastnames_and_id

    def get_info_lastname(self, dct, daterange ):
        """
        
        Возвращает список словарей формата
        [
            {
                'average_time': 16000.0,  --среднее время звонка одного оператора за день в милисекундах 
                'duration_call': 16000, --сумма всех звонков оператора за день в миллисекундах
                'name_employee': 'Мария Золотарева [Маткласс]', -- фио оператора по тегу Сотрудник
                'the_date': '15.11.2022', -- день совершения звонков
                'time_call': [1668528117], -- отсортированный список список начал всех unixstapm
                'time_first': 1668528117, -- начало первого unixstapm
                'time_last': 1668528117, -- начало последнего unixstapm
                'total': 1, -- количество звонков за день
                'work_time': 16.0 -- время работы в секундах
            },
            {

            },
            {
            },
        ]

        """
        data_copy=[]
        for key, value in dct.items():
            data={}
            time_call=[]
            duration_call=0
            total=0
            last_duration=0
            for item in value:
                res = Req().get( f"{self._MAIN_URL}call/"+item )
                calls= json.loads(res.text)
                time_call.append(calls['call_time'])
                duration_call=duration_call+calls['duration']
                total = total+1
                if item == value[-1] : last_duration=calls['duration']/1000
            time_call=sorted(time_call)
            data['name_employee']=key
            data['time_first']= time_call[0]
            data['time_last']=time_call[-1]
            data['work_time']=time_call[-1]-time_call[0]+last_duration
            data['time_call']=time_call
            data['duration_call']=duration_call
            data['total']=total
            data['average_time']= duration_call/total
            data['the_date']=daterange
            data_copy.append(data)
        return data_copy
    
    def output_results(self, data, daterange ):
        """

        Выводит значения из get_info_lastname в json файл в том же формате
        
        """
        with open('output_of_result-'+daterange+'.json', 'w') as file:
            json.dump(data, file)



if __name__=="__main__":
    pass