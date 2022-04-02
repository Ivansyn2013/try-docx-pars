import docx
import json
import time, datetime

import re

temp_date = datetime.datetime.now()
RE_DATE = re.compile(r'(\d{2}\.\d{2})\.\d{4}')


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


START_DATE = datetime.datetime.strptime(get_config_file().get('даты проведения цикла'), '%d.%m.%Y')
# забор файла
doc = docx.Document(r'G:\Python project\try-docx-pars\sample_files\PK(144hours).docx')

# print(len(doc.paragraphs))

# print((doc.paragraphs[4].text))

line = doc.paragraphs[4]
list = []

# for el1 in doc.paragraphs:
#     for el2 in el1.runs:
#         print(el2.text.split(' ')[0])

# print(line.runs[0].text)

# выбор первого столбца таблицы

tabel = doc.tables[0]
print(len(tabel.columns[0].cells))

# print(tabel.columns[0].cells[4].text)
# print(tabel.cell(4,0).text)
# print(tabel.cell(0,0).text)
k = 0

i = 0

for tabel in doc.tables:

    for cell in tabel.columns[0].cells:

        if time_valid(cell.text) and i < 1:
            cell.text = re.sub(RE_DATE, (START_DATE.date().strftime('%d.%m.%Y')), cell.text)

            i += 1
        elif time_valid(cell.text):
            if i % 3 == 0:
                delta_time = START_DATE + datetime.timedelta(days=i // 3)
                if delta_time.weekday() == 6:
                    delta_time += datetime.timedelta(days=1)

                cell.text = re.sub(RE_DATE, delta_time.date().strftime('%d.%m.%Y'), cell.text)
                cell.text.replace(cell.text, delta_time.date().strftime('%d.%m.%Y'))
                print(cell.text)
            else:
                pass
        i += 1
        # print(cell.text)

# сохранение объекта файла
doc.save(r'G:\Python project\try-docx-pars\ready1.docx')
