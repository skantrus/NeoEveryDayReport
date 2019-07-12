import pandas as pn, openpyxl, datetime
from dateutil import parser
import os
import supporting_scripts

# supporting_scripts.clear_current_report(current_dir,'Report.xlsx') # Our every day report   line 14
# OSReport = 'Отчет по ОС (ВТБ. Управление заявками ИТ).xls'  # Filename of Every day Report line 23
# OSExpired = pn.read_html(current_dir + 'Просрочки (ВТБ. Управление заявками ИТ).xls')[1]  # 1st imported file from Bank Jira   line 145
# Report.save(current_dir + 'testt.xlsx')  # Save report to file xxx   line 225

def main():
    current_dir='E:\работа\_отчёты\_ежедневный\\'##Path to files

    supporting_scripts.clear_current_report(current_dir,'Отчет по OS.xlsx')#current_dir+filename or 'E:\\_proj\\eoflex\\_everyday\\Report.xlsx'

    curdate=parser.parse('2019-07-11').date()
    os_control,omni_list,alm_list=supporting_scripts.import_from_google_tables_oscontrol(curdate)
    write_data_to_osreport(current_dir, os_control, omni_list, alm_list)

    input()

    
def write_data_to_osreport(current_dir,os_control,omni_list,alm_list):

    def Count1():
        """Function to write data into a OSReport.Count1"""
        OSReport = pn.read_html(current_dir+'Отчет по OS (ВТБ. Управление заявками ИТ).xls')[1]  # pandas.core.frame.DataFrame
        OSReport = list(OSReport.get_values())  # to list

        Counting1=Report['Подсчет1']
        for NumRow in range(len(OSReport)):
            for NumCell in range(len(OSReport[NumRow])):
                Counting1.cell(row=NumRow + 2, column=NumCell + 2).value = OSReport[NumRow][NumCell]
        return

    def table_for_count2(oscontrol):
        oscontrol.reverse()
        table_for_count2=Report['Таблица для Подсчета2']

        for index in range(len(oscontrol)):
            try:
                table_for_count2.cell(row=index + 2, column=1).value = oscontrol[index][4]
                table_for_count2.cell(row=index + 2, column=2).value = oscontrol[index][5]
                table_for_count2.cell(row=index + 2, column=3).value = oscontrol[index][6]
                table_for_count2.cell(row=index + 2, column=4).value = parser.parse(oscontrol[index][1], dayfirst=True)
                table_for_count2.cell(row=index + 2, column=4).number_format = 'DD.MM.YY H:MM;@'
                table_for_count2.cell(row=index + 2, column=5).value = parser.parse(oscontrol[index][0], dayfirst=True)
                table_for_count2.cell(row=index + 2, column=5).number_format = 'DD.MM.YY H:MM;@'
                table_for_count2.cell(row=index + 2, column=6).value = parser.parse(oscontrol[index][2], dayfirst=True)
                table_for_count2.cell(row=index + 2, column=6).number_format = 'DD.MM.YY H:MM;@'
                table_for_count2.cell(row=index + 2, column=7).value = oscontrol[index][7]
                table_for_count2.cell(row=index + 2, column=8).value = oscontrol[index][8]
                table_for_count2.cell(row=index + 2, column=9).value = oscontrol[index][9]
                table_for_count2.cell(row=index + 2, column=10).value = oscontrol[index][10]
                table_for_count2.cell(row=index + 2, column=11).value = oscontrol[index][11]
            except Exception as e:
                print('Некорретное значение в Отчёте, страница "Таблица для Подсчета2", строка',index+2,'/nНеобходимо исправить в OS control и перезапустить скрипт')
                continue
        return

    def table_for_cont_alm(almlist):
        almlist.reverse()
        table_for_count_alm=Report['Таблица для Посчета ALM']

        for index in range(len(almlist)):
            try:
                table_for_count_alm.cell(row=index + 2, column=1).value = parser.parse(almlist[index][0], dayfirst=True)
                table_for_count_alm.cell(row=index + 2, column=1).number_format = 'DD.MM.YY H:MM;@'
                table_for_count_alm.cell(row=index + 2, column=2).value = parser.parse(almlist[index][1], dayfirst=True)
                table_for_count_alm.cell(row=index + 2, column=2).number_format = 'DD.MM.YY H:MM;@'
                table_for_count_alm.cell(row=index + 2, column=3).value = parser.parse(almlist[index][2], dayfirst=True)
                table_for_count_alm.cell(row=index + 2, column=3).number_format = 'DD.MM.YY H:MM;@'
                table_for_count_alm.cell(row=index + 2, column=4).value = almlist[index][3]
                table_for_count_alm.cell(row=index + 2, column=5).value = almlist[index][4]
                table_for_count_alm.cell(row=index + 2, column=6).value = almlist[index][5]
                table_for_count_alm.cell(row=index + 2, column=7).value = almlist[index][6]
                table_for_count_alm.cell(row=index + 2, column=8).value = almlist[index][7]
                table_for_count_alm.cell(row=index + 2, column=9).value = almlist[index][8]
                table_for_count_alm.cell(row=index + 2, column=10).value = almlist[index][9]
                table_for_count_alm.cell(row=index + 2, column=11).value = almlist[index][10]
                table_for_count_alm.cell(row=index + 2, column=12).value = almlist[index][11]
            except Exception as e:
                print('Некорретное значение в Отчёте, страница "Таблица для Посчета ALM", строка',index+2,'/nНеобходимо исправить в ALM control и перезапустить скрипт')
                continue

    def omni(omnilist):
        omnilist.reverse()
        omni=Report['Таблица для Подсчета Omni']
        for i in range(len(omnilist)):
            try:
                omni.cell(row=i + 2, column=1).value = omnilist[i][0]
                omni.cell(row=i + 2, column=2).value = parser.parse(omnilist[i][1], dayfirst=True)
                omni.cell(row=i + 2, column=2).number_format = 'DD.MM.YY H:MM;@'
                omni.cell(row=i + 2, column=3).value = omnilist[i][2]
                omni.cell(row=i + 2, column=4).value = omnilist[i][3]
                omni.cell(row=i + 2, column=5).value = parser.parse(omnilist[i][4], dayfirst=True)
                omni.cell(row=i + 2, column=5).number_format = 'DD.MM.YY H:MM;@'
                omni.cell(row=i + 2, column=6).value = omnilist[i][5]
                omni.cell(row=i + 2, column=7).value = omnilist[i][6]
                omni.cell(row=i + 2, column=8).value = omnilist[i][7]
                omni.cell(row=i + 2, column=9).value = omnilist[i][8]
                omni.cell(row=i + 2, column=10).value = omnilist[i][9]
                omni.cell(row=i + 2, column=11).value = omnilist[i][10]
            except Exception as e:
                print('Некорретное значение в Отчёте, страница "Таблица для Подсчета Omni", строка',index+2,'/nНеобходимо исправить в Реестр обращений OMNI и перезапустить скрипт')
                continue

    def return_correct_system(value):
        if value == 'УСБС-Middle-СЗП: Неофлекс':
            return 'УСБС-Middle-СЗП'
        elif value == 'УСБС-Front: Неофлекс':
            return 'УСБС-Фронт'
        elif value == 'УСБС-Middle: Гарантированная поддержка' or value == 'УСБС-Middle: Неофлекс':
            return 'УСБС-Middle'
        elif value == 'PLM':
            return 'PLM'
        elif value == 'MDS-ULBS':
            return 'MDS-Фронт'
        elif value == 'Not bound':
            return 'MDS-Middle'
        elif value == 'TS 2.0: Неофлекс':
            return 'TC2.0'
        elif value == 'Not bound':
            return 'KK'
        else:
            return ''







    Report = openpyxl.load_workbook(current_dir+'Отчет по OS.xlsx')  # Our Everyday Report

    Count1()#put data into Report.Count1
    table_for_count2(os_control)#..
    omni(omni_list)#..
    table_for_cont_alm(alm_list)#put data into Report.table_for_cont_alm

    ##Start of input Expired OSes##

    ##Expired Oses in 'Просрочки (ВТБ. Управление заявками ИТ)'
    OSExpired = pn.read_html(current_dir+'Просрочки_1 (ВТБ. Управление заявками ИТ).xls')[1]
    OSExpired = list(OSExpired.get_values())
    oses_expired_probably_new_set = set()
    oses_expired_probably_new=dict()
    for row in OSExpired:
        oses_expired_probably_new_set.add(row[1])
        oses_expired_probably_new.update({str(row[1]):[row[1],row[2],row[6],row[0],parser.parse(row[7], dayfirst=True)]})

    ##Expired Oses already in Report
    Expired=Report["Просрочки"]
    itr=16
    oses_expired_probably_set=set()
    oses_expired_probably=dict()
    while True:
        if Expired.cell(row=itr, column=1).value != None:
            oses_expired_probably_set.add(str(Expired.cell(row=itr, column=1).value))
            row_in_Expired=[]
            for x in range(1,8):
                row_in_Expired.append(str(Expired.cell(row=itr, column=x).value))              #str(oses_expired_probably.update({str(Expired.cell(row=i, column=x))
                Expired.cell(row=itr, column=x).value = None
                Expired.cell(row=itr, column=x).border.bottom.border_style=None
                Expired.cell(row=itr, column=x).border.bottom.style = None
                Expired.cell(row=itr, column=x).border.left.border_style = None
                Expired.cell(row=itr, column=x).border.left.style = None
                Expired.cell(row=itr, column=x).border.right.border_style = None
                Expired.cell(row=itr, column=x).border.right.style = None
                Expired.cell(row=itr, column=x).border.top.border_style = None
                Expired.cell(row=itr, column=x).border.top.style = None
            row_in_Expired[4]=parser.parse(row_in_Expired[4],dayfirst=True)
            oses_expired_probably.update({str(row_in_Expired[0]):row_in_Expired})

            itr+=1
            continue
        else:
            break


    oses_union=oses_expired_probably_new_set.intersection(oses_expired_probably_set)
    oses_expired_probably_new_set.difference_update(oses_union)
    i=16
##    while True:
    for new_one_row in oses_expired_probably:
        if new_one_row[0] in oses_union:
            for x in range(1, 8):
                Expired.cell(row=i, column=x).value = new_one_row[x - 1]
                Expired.cell(row=i, column=x).border.bottom.border_style = 'thin'
                Expired.cell(row=i, column=x).border.bottom.style = 'thin'
                Expired.cell(row=i, column=x).border.left.border_style = 'thin'
                Expired.cell(row=i, column=x).border.left.style = 'thin'
                Expired.cell(row=i, column=x).border.right.border_style = 'thin'
                Expired.cell(row=i, column=x).border.right.style = 'thin'
                Expired.cell(row=i, column=x).border.top.border_style = 'thin'
                Expired.cell(row=i, column=x).border.top.style = 'thin'
        i += 1
##        
##
##
##            
##        try:
##            new_one_os=oses_union.pop()
##            new_one_row = oses_expired_probably[new_one_os]
##            for x in range(1, 8):
##                Expired.cell(row=i, column=x).value = new_one_row[x - 1]
##                Expired.cell(row=i, column=x).border.bottom.border_style = 'thin'
##                Expired.cell(row=i, column=x).border.bottom.style = 'thin'
##                Expired.cell(row=i, column=x).border.left.border_style = 'thin'
##                Expired.cell(row=i, column=x).border.left.style = 'thin'
##                Expired.cell(row=i, column=x).border.right.border_style = 'thin'
##                Expired.cell(row=i, column=x).border.right.style = 'thin'
##                Expired.cell(row=i, column=x).border.top.border_style = 'thin'
##                Expired.cell(row=i, column=x).border.top.style = 'thin'
##            i += 1
##            continue
##        except KeyError:
##            try:
    while True:
        new_one_os=oses_expired_probably_new_set.pop()
        nne_row=oses_expired_probably_new[new_one_os]
        new_one_row[2] = return_correct_system(new_one_row[2])
        for x in range(1, 6):
            Expired.cell(row=i, column=x).value = new_one_row[x - 1]
            Expired.cell(row=i, column=x).border.bottom.border_style = 'thin'
            Expired.cell(row=i, column=x).border.bottom.style = 'thin'
            Expired.cell(row=i, column=x).border.left.border_style = 'thin'
            Expired.cell(row=i, column=x).border.left.style = 'thin'
            Expired.cell(row=i, column=x).border.right.border_style = 'thin'
            Expired.cell(row=i, column=x).border.right.style = 'thin'
            Expired.cell(row=i, column=x).border.top.border_style = 'thin'
            Expired.cell(row=i, column=x).border.top.style = 'thin'
        flag=1
        i += 1
    
##    except KeyError:
##        break
    try:
        print('В отчёт добавлены новые OS в просрочке\nНеобходимо проверить дату,выставить резолюцию по просрочке и причину\n\n') if flag else print()
    except Exception:
        print('Новые просрочки отсутствуют\n\n')
    Report.save(current_dir+'Отчет по OS new.xlsx')
    ##End of input Expired OSes##


if __name__ == '__main__':
    main()




