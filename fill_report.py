import pandas as pn, openpyxl, datetime
from dateutil import parser
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import supporting_scripts


def main():
    os_control,omni_list,alm_list=import_from_google_tables_oscontrol()
    write_data_to_osreport(os_control,omni_list,alm_list)



def write_data_to_osreport(os_control,omni_list,alm_list):

    def Count1():
        """Function to write data into a OSReport.Count1"""
        OSReport = pn.read_html('E:\\_proj\\Neoflex\\_everyday\\OSReport.xls')[1]  # pandas.core.frame.DataFrame
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
        if value in 'УСБС-Middle-СЗП: Неофлекс':
            return 'УСБС-Middle-СЗП'
        elif value in 'УСБС-Front: Неофлекс':
            return 'УСБС-Фронт'
        elif value in 'УСБС-Middle: Гарантированная поддержка' or value in 'УСБС-Middle: Неофлекс':
            return 'УСБС-Middle'
        elif value in 'PLM':
            return 'PLM'
        elif value in 'MDS-ULBS':
            return 'MDS-Фронт'
        elif value in 'Not bound':
            return 'MDS-Middle'
        elif value in 'TS 2.0: Неофлекс':
            return 'TC2.0'
        elif value in 'Not bound':
            return 'KK'
        else:
            return ''






    supporting_scripts.clear_current_report()
    Report = openpyxl.load_workbook('E:\\_proj\\Neoflex\\_everyday\\Report.xlsx')  # Our Everyday Report

    Count1()
    table_for_count2(os_control)
    omni(omni_list)
    table_for_cont_alm(alm_list)

    ##Expired Oses
    OSExpired = pn.read_html('E:\\_proj\\Neoflex\\_everyday\\Просрочки (ВТБ. Управление заявками ИТ) (3).xls')[1]
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
    while True:
        try:
            new_one_os=oses_union.pop()
            new_one_row = oses_expired_probably[new_one_os]
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
            continue
        except KeyError:
            try:
                new_one_os=oses_expired_probably_new_set.pop()
                new_one_row=oses_expired_probably_new[new_one_os]
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
                continue
            except KeyError:
                break

    try:
        print('В отчёт добавлены новые OS в просрочке\nНеобходимо проверить дату,выставить резолюцию по просрочке и причину\n\n') if flag else print()
    except Exception:
        print('Новые просрочки отсутствуют\n\n')
    Report.save('E:\\_proj\\Neoflex\\_everyday\\testt.xlsx')  # E:\\_proj\\Neoflex\\_everyday\\Report.xlsx

def import_from_google_tables_oscontrol():
    """Get Data from Google Tables"""
    omni_key,omni_sheet_name = '1drbPbjMKGbODn1FGqmR0VdRzh3zyaxGVjq1prTAu2rs','2019'
    oscontrol_key,oscontrol_sheet_name='1HiDdPqB_-ro4Iu0RplDnt3k8lEjDd_UiOBPKvyWOPuY','Сводная таблица по OS'
    almcontrol_key,almcontrol_sheet_name='1SDSEhgtQTHR9a69BfYE1Fd2DnMH6fIFBEsvyHldwmM8','Сводная таблица по ALM'####for test '13criem2KpgGtQA3fjv2BxH3f3-r3WWi5TJqod4bAQfg','Лист1'

    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('NeoflexReports-376d91d0718a.json', scope)
    client = gspread.authorize(creds)
    curdate = parser.parse('2019-07-04').date()  ##should be =datetime.datetime.now().date()

    oscontrol_sheet = client.open_by_key(oscontrol_key).worksheet(oscontrol_sheet_name).get_all_values()  ##Should be *.worksheet('OSCONTROL')
    oscontrol_list = []  # list with OS with all filled cells
    for i in range(len(oscontrol_sheet), 0, -1):
        if oscontrol_sheet[i - 1][0] != '' and oscontrol_sheet[i - 1][1] != '' and oscontrol_sheet[i - 1][2] != '' and oscontrol_sheet[i - 1][11] != '':
            ##Currenct day from 00:00 to 23:59
            try:
                if parser.parse(oscontrol_sheet[i - 1][2], dayfirst=True).date() == curdate:
                    flag1 = 1
                    oscontrol_list.append(oscontrol_sheet[i - 1][0:12])
                else:
                    continue
            except Exception as e:
                print(str(e),oscontrol_sheet[i - 1],i)
                continue
        else:
            try:
                if flag1:  # if flag exist - we went through current date
                    break
                else:
                    continue  # if flag does not exist - we still dont get current date - WE ARE IN THE FUTURE!
            except UnboundLocalError:
                continue

    omni_sheet = client.open_by_key(omni_key).worksheet(omni_sheet_name).get_all_values()
    omni_list=[]
    for i in range(len(omni_sheet), 0, -1):
        if omni_sheet[i - 1][5] != '' and omni_sheet[i - 1][0] != '' and omni_sheet[i - 1][1] != '' and omni_sheet[i - 1][2] != '':
            try:
                if parser.parse(omni_sheet[i - 1][5], dayfirst=True).date() == curdate:  ##should be ==curdate
                    flag2 = 1
                    omni_list.append(omni_sheet[i - 1][0:2]+omni_sheet[i - 1][3:6]+omni_sheet[i - 1][7:])
            except Exception as e:
                print(str(e),omni_sheet[i - 1],i)
                continue
        else:
            try:
                if flag2:  # if flag exist - we went through current date
                    break
                else:
                    continue  # if flag does not exist - we still dont get current date - WE ARE IN THE FUTURE!
            except UnboundLocalError:
                continue

    alm_sheet= client.open_by_key(almcontrol_key).worksheet(almcontrol_sheet_name).get_all_values()
    alm_list=[]
    for i in range(len(alm_sheet), 0, -1):
        if alm_sheet[i - 1][0] != '' and alm_sheet[i - 1][1] != '' and alm_sheet[i - 1][2] != '' and alm_sheet[i - 1][11] != '':
            ##Currenct day from 00:00 to 23:59
            try:
                if parser.parse(alm_sheet[i - 1][2], dayfirst=True).date() == curdate:
                    flag3 = 1
                    alm_list.append(alm_sheet[i - 1][0:12])
                else:
                    continue
            except Exception as e:
                print(str(e),alm_sheet[i - 1],i)
                continue
        else:
            try:
                if flag3:  # if flag exist - we went through current date
                    break
                else:
                    continue  # if flag does not exist - we still dont get current date - WE ARE IN THE FUTURE!
            except UnboundLocalError:
                continue


    return oscontrol_list,omni_list,alm_list

if __name__ == '__main__':
    main()




