import docx
import re



RE_LINESTART = re.compile(r'^\d{2} *')
bibl = docx.Document(r'C:\Users\Александр\PycharmProjects\try-docx-pars\sample_files\litspisok.docx')
i = 0
origin_list = []
cell_list = []
# количество абзацев в документе
print(len(bibl.paragraphs))
print(len(bibl.tables))
for par in bibl.paragraphs:
    print(par.text)

table = bibl.tables[0].rows
for row in table:
    i += 1
    cell_list = []
    for cell in row.cells:
        cell_list.append(cell.text)

    origin_list.append(cell_list)

# сортировка по второму эдементу списка
origin_list.sort(key=lambda x: x[1])
#print(origin_list)
# создание сортированого списка.
for inx, el in enumerate(origin_list, 1):
    el.append(inx)
#print(origin_list)
# создание словаря со значения бывшего и нового порядка
origins_dict = {y: [x, z] for x, y, z in tuple(origin_list)}

print(origins_dict)

# текст первого абзаца в документе
# print(oks_doc12.paragraphs[0].text)

# текст второго абзаца в документе
# print(oks_doc12.paragraphs[6].text)
# print(oks_doc12.paragraphs[7].text)
# print(oks_doc12.paragraphs[8].text)
# print(oks_doc12.paragraphs[9].text)
# текст первого Run второго абзаца
