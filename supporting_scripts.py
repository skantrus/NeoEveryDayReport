import pandas as pn, openpyxl
from copy import copy

def clear_current_report():
    Report = openpyxl.load_workbook('E:\\_proj\\Neoflex\\_everyday\\Report.xlsx')  # Our Everyday Report

    counting1 = Report['Подсчет1']
    index=2
    while True:
        if counting1.cell(row=index, column=2).value != None:
            for i in range(2,9):
                counting1.cell(row=index, column=i).value = None
        else:
            break
        index += 1

    table_for_count2 = Report['Таблица для Подсчета2'] #table_for_count_alm.cell(row=2, column=1)
    index = 2                                          #table_for_count2.cell(row=2, column=1)
    while True:
        if table_for_count2.cell(row=index, column=1).value != None:
            for i in range(1,12):
                table_for_count2.cell(row=index, column=i).value = None
        else:
            break
        index+=1

    table_for_count_alm = Report['Таблица для Посчета ALM']
    index = 2
    while True:
        if table_for_count_alm.cell(row=index, column=1).value != None:
            for i in range(1,13):
                table_for_count_alm.cell(row=index, column=i).value = None

                # new_cell.font = copy(cell.font)
                # new_cell.border = copy(cell.border)
                # new_cell.fill = copy(cell.fill)
                # new_cell.number_format = copy(cell.number_format)
                # new_cell.protection = copy(cell.protection)
                # new_cell.alignment = copy(cell.alignment)
                table_for_count_alm.cell(row=index, column=i).font = copy(table_for_count_alm.cell(row=2, column=16).font)
                table_for_count_alm.cell(row=index, column=i).border = copy(table_for_count_alm.cell(row=2, column=16).border)
                table_for_count_alm.cell(row=index, column=i).fill = copy(table_for_count_alm.cell(row=2, column=16).fill)

        else:
            break
        index+=1

    omni = Report['Таблица для Подсчета Omni']
    index = 2
    while True:
        if omni.cell(row=index, column=1).value != None:
            for i in range(1, 12):
                omni.cell(row=index, column=i).value = None
                omni.cell(row=index, column=i).font = copy(omni.cell(row=2, column=16).font)
                omni.cell(row=index, column=i).border = copy(omni.cell(row=2, column=16).border)
                omni.cell(row=index, column=i).fill = copy(omni.cell(row=2, column=16).fill)
        else:
            break
        index += 1

    Report.save('E:\\_proj\\Neoflex\\_everyday\\Report.xlsx')  # E:\\_proj\\Neoflex\\_everyday\\Report.xlsx
