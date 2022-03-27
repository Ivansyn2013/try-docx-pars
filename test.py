import docx
import json
import time, datetime
from itertools import groupby

import re

temp_date = datetime.datetime.now()
RE_DATE = re.compile(r'(\d{2}\.\d{2}).\d{4}')


def run_config_file():
    '''создание файла для конфигурации'''
    config_dict = {'название цикла': '', 'даты проведения цикла': 'fdf', 'количество часов': ''}

    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config_dict, f, ensure_ascii=False, indent=4)
    return None


def get_config_file():
    '''получение конфишурации из файла'''
    with open('config.json', 'r', encoding='utf-8') as f:
        config_dict = json.load(f)

    return config_dict


def time_valid(date_str):
    import re

    RE_DATE = re.compile(r'(\d{2}\.\d{2}).\d{4}')
    if RE_DATE.search(date_str) != None:
        return True
    else:
        return False

doc = docx.Document(r'G:\Python project\try-docx-pars\sample_files\PP(572hours).docx')




#doc = docx.Document(r'G:\Python project\try-selenium\sample_files\PK(144hours).docx')

#tabel = doc.tables[0]

# datet = datetime.datetime.strptime('02.01.2022','%d.%m.%Y')
#
#
# #print(help(datet.date().strftime('%d.%m.%Y')))
# print(datet.date().strftime('%d.%m.%Y'))
# print(datet.weekday())
