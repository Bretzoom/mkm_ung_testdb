import os
import pandas as pd
import openpyxl as oxl

listdir = os.listdir()

if "result.xlsx" in listdir:
    print("Хотите сохранить предыдущий результат? Да/Нет")
    save = str(input())
    if save == "Да" or save == "да" or save == "ДА":
        if "result_backup.xlsx" in listdir:
            print("Предыдущий сохраненный результат будет удален. Продолжить? Да/Нет")
            delete = str(input())
            if delete == "Да" or delete == "да" or delete == "ДА":
                os.remove("result_backup.xlsx")
                os.rename('result.xlsx', 'result_backup.xlsx')
        else:
            os.rename("result.xlsx", "result_backup.xlsx")
    else:
        os.remove('result.xlsx')
        df = pd.DataFrame()
        df.to_excel("result.xlsx")
else:
    df = pd.DataFrame()
    df.to_excel("result.xlsx")
oilfields = []
places = []
wells = []
dates = []
temps = []
amount = []
gellants = []
crosslinx = []
destructors = []
sources = []
h_links = []
counter = int(input('Введите номер последней введенной в таблицу записи:'))
trigger = 0

for item in listdir:
    temp = item.split(" ")
    if "database.py" not in temp and "database.exe" not in temp and "result.xlsx" not in temp and "result_backup.xlsx" not in temp and ".idea" not in temp and '~$result.xlsx' not in temp:
        counter += 1
        h_links.append('=ГИПЕРССЫЛКА("' + str(os.path.abspath(item)).replace(" ", "%20") + '","' + str(counter) + '")')
        dates.append(temp[0])
        oilfields.append(temp[1])
        places.append(temp[3])
        wells.append(temp[5])
        temps.append(temp[6])
        amount.append(temp[7])
        os.chdir(item)
        listdir2 = os.listdir()
        for item2 in listdir2:
            if '.xlsm' in item2 or '.xlsx' in item2:
                try:
                    excel = pd.read_excel(os.getcwd() + '\\' + item2, 'Данные')
                    sources.append(excel.iloc[4][2])
                    if '/' in excel.iloc[20][4]:
                        gellants.append(excel.iloc[20][4].split('/')[0])
                    else:
                        gellants.append(excel.iloc[20][4])
                    if '/' in excel.iloc[35][4]:
                        crosslinx.append(excel.iloc[35][4].split('/')[0])
                    else:
                        crosslinx.append(excel.iloc[35][4])
                    destructors.append(excel.iloc[43][4].split('-')[0])
                    listdir2 = []
                except PermissionError:
                    print('--- Oшибка! Вы забыли закрыть excel фаил ---')
        os.chdir('../')
        trigger = 0
    print(temp)
resultlist = [h_links, oilfields, places, wells, sources, temps, gellants, crosslinx, destructors, dates, amount]
try:
    excel = pd.read_excel(r"result.xlsx", index_col=0)
    row = 0
    wb = oxl.load_workbook(r"result.xlsx")
    sheet = wb.active
    columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'J']
    for i in range(0, len(columns) - 1):
        for j in range(0, len(oilfields)):
            sheet[columns[i] + str(j + 1)] = str(resultlist[i][j])
    wb.save(r"result.xlsx")
except PermissionError:
    print('--- Oшибка! Вы забыли закрыть excel фаил ---')
