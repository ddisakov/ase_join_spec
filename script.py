#Скрипт для выполнения объединения спецификаций (без кабельных журналов)
#Первым делом выявляет отсутствие спецификаций
# Файл должен находиться в определённом месте


#19/02/2023
# Алгоритм использования:
# Распаковываем все архивы - 1 проход, потом архиватор комментируем, а то заебёт
# Заходим в файл no_data. Там будут спецификации, на которые нет рабочих файлов. Запрашиваем - пихаем по папкам
# Открываем файл change_name_sheets и идём по спецификациям переименовыватть листы на CommonList
# Файл revisions пока вручную, нужно писать регулярку на проверку открытия последней версии. Поэтому шерстим, что он не накопировал лишнего по этим файлам
# Файл paths - пути, которые скрипт прошёл для сверки, дебага



import re
import pandas as pd
import numpy as np
import os
import csv
import patoolib
import openpyxl

#вывод данных в csv
def to_csv(lst, file_name):

    with open(file_name, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        for item in lst:
            csv_writer.writerow([item])
    

#Извлечение номеров спецификаций по введённому пользователем объекту или объектам
to_do_list = 'list.xlsx'
doc_list = pd.read_excel(to_do_list, sheet_name=0)

user_object = '03UYX'
#user_object = input('Введите KKS объекта \n').upper()

#Отсеиваем кабельные журналы, оставляем только спецификации Если их придётся делать, то надо будет писать под них
df_filtered = doc_list[(doc_list['Здание'].str.contains(user_object)) & (doc_list['Ответсвенный ЗД/ДСАР'].str.contains('Кунаш')) & (doc_list['Шифр ВОР/код спецификации'].str.endswith('S0001')) & (doc_list['Шифр ВОР/код спецификации'].str.contains('.MB0001') == False)]
user_specifications = df_filtered['Шифр ВОР/код спецификации']

#Формирование путей для посещения
visit=[]
for i in user_specifications:
    for root, dirs, filenames in os.walk('../Projects/'+user_object):
        if i in root:
            visit.append(root)
            #Разархиватор, включать только первый раз, если неохота доставать данные. Важно оставить под условием иначе будт распаковывать все папки
            for k in filenames:
                if k.upper().endswith('.RAR'):
                    patoolib.extract_archive(root+'/'+k, outdir=root)
                    


nodata, revisions, wrong_names_sheets, paths = [], [], [], []

#Проверка на наличие рабочих файлов и других ревизий
for i in visit:
    if '=C01' not in i:
        revisions.append(i)
    flag = 0
    a = os.listdir(i)
    for j in a:
        if j.lower().endswith('.xls') or j.lower().endswith('.xlsx') or j.lower().endswith('.xlsm'):
            flag = 1
    if flag == 0:
        nodata.append(i)

#Запись в CSV спецификаций, у которх есть ревизии и в которых нет рабочих файлов
to_csv(nodata, 'no_data.csv')
to_csv(revisions, 'revisions.csv')
# Удаляем из спика директории, где нет рабочки
for i in nodata:
    visit.remove(i)




db_result = pd.read_excel('Template.xlsx')
db_result.columns=[1,2,3,4,5,6,7,8,9,10,11,12,13,14]
#Шаблон для спецификации
template = r'RPR\.\d{4}\.\d{2}[A-Z]{3}\.\S{1,3}\.[A-Z]{2}\.[A-Z]{2}\d{4}\.[A-Z]\d{4}\=C0\d'


#Доступ к данным. Везде листы названы по разному, придётся потыкать и вручную переименовывать листы. Список таких листов в csv
for i in visit:
    spec_num = re.findall(template, i)         
    for j in os.listdir(i):
        curr_path = i + '/' + j
        if j.lower().endswith('.xls') or j.lower().endswith('.xlsx') or j.lower().endswith('.xlsm'):
            try:
                
                data = pd.read_excel(curr_path, sheet_name='CommonList', header=2)
                data.columns = [1,2,3,4,5,6,7,8,9]
                data = data.dropna(subset=[6])
                #Формирование датафрейма такого же формата и добавление к образцу
                data2 = pd.DataFrame({1: [spec_num[0][:-4] for i in range(len(data))],
                                      2: [spec_num[0][-3:] for i in range(len(data))],
                                      3: data[1],
                                      4: data[2],
                                      5: data[2],
                                      6: data[3],
                                      7: [None for i in range(len(data))],
                                      8: [None for i in range(len(data))],
                                      9: data[5],
                                      10:[None for i in range(len(data))],
                                      11: data[6],
                                      12: data[7],
                                      13: data[8],
                                      14: data[9]
                                      })
                db_result = pd.concat([db_result, data2])
            except ValueError:
                wrong_names_sheets.append(j)
            finally:
                paths.append(curr_path)
            
to_csv(visit, 'visit.csv')
to_csv(wrong_names_sheets, 'change_name_sheet.csv')
to_csv(paths, 'paths.csv')

db_result.to_excel('keks.xlsx')                
            