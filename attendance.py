#!/usr/bin/python
# -*- codingLutf-8 -*-
import os
import xlwings as xw
from shutil import copyfile, move
from sys import exit, exc_info
from datetime import datetime

def copyonetemplate(srcfile, dstfile):
    '''
    拷贝模板文件到指定目录
    :param srcfile:
    :param dstfile:
    :return:
    '''
    try:
        copyfile(srcfile, dstfile)
    except IOError as e:
        print("Unable to Copy file. %s"%e)
        exit(1)
    except:
        print("Unexpected error:", exc_info())
        exit(1)

def moveoneresultfile(path, filename):
    '''

    :param path:
    :param filename:
    :return:
    '''
    try:
        move(filename, path)
    except IOError as e:
        print("Unable to move file. %s"%e)
        exit(1)
    except:
        print("Unexpected error:", exc_info())
        exit(1)

def getindexfdate(workdate):
    return workdate.day + 2

def setcuryearmonth(srcsheet, date):
    '''
    设置当前月和年
    :param srcsheet:
    :param date:
    :return:
    '''
    srcsheet.range('B4').value = str(date.month) + '月'
    srcsheet.range('AH4').value = date.year
    return

def addnewperson(srcsheet, personname):
    '''
    新增一个员工
    :param srcsheet:
    :param personname:
    :return:
    '''
    srcsheet.api.Rows(7).Insert()
    srcsheet.range(7, 2).value = personname
    #设置员工行高度
    srcsheet.range(7, 2).row_height = 50
    return

def setbackgrand(selete, moringtime, afternoontime):
    if moringtime.hour >= 9 and moringtime.minute != 0:
        #设置迟到
        selete.color = (222, 116, 101)
    elif afternoontime.hour <= 17 and afternoontime.minute < 30:
        #设置早退
        selete.color = (244, 205, 172)
    else:
        #设置正常
        selete.color = (196, 208, 157)
    return

def addnewattence(srcsheet, morningtime, afternoontime):
    '''
    添加一天的出勤
    :param srcsheet: 需要设置的表
    :param morningtime: 上班时间
    :param afternoontime: 下班时间
    :return:
    '''
    #获取上班时间列
    cols = getindexfdate(morningtime)
    selete = srcsheet.range(7, cols)
    selete.value = \
        "{mtime}\r\n~{atime}".format(mtime=morningtime.strftime('%H:%M:%S'), atime=afternoontime.strftime('%H:%M:%S'))
    #根据上下班时间设置背景颜色
    setbackgrand(selete, morningtime, afternoontime)
    return

def core(app, srcfile):
    copyonetemplate('template.xlsx', r'员工考勤时间表.xlsx')
    recodewb = app.books.open(r'员工考勤时间表.xlsx')
    srcwb = app.books.open(srcfile)
    srcsheet = srcwb.sheets[0]
    recordsheet = recodewb.sheets('考勤')
    cols = srcsheet.used_range.shape[0]
    lastname = ""
    moningdatetime = None
    afternoondatetime = None
    indexdate = None
    hasSetMonth = False
    for x in range(2, cols):
        name = srcsheet.range(x, 1).value
        if (name is None) or (str(name).isspace() is True):
            break
        srcsheet.range(x, 4).formula = "=B{row}+C{row}".format(row=x)
        #读取日期及时间
        recordtime = srcsheet.range(x, 4).value
        if not isinstance(recordtime, datetime):
            continue

        #刷新表格是年及月份
        if not hasSetMonth:
            setcuryearmonth(recordsheet, recordtime);
            hasSetMonth = True

        #当人员发生变化时，需要新增一列
        if lastname != name:
            #人员变化后， 需要将最后一组数据填入表格
            if lastname != "":
                addnewattence(recordsheet, moningdatetime, afternoondatetime)
            lastname = name
            addnewperson(recordsheet, name)
            # 切换游标时间
            indexdate = recordtime
            moningdatetime = None
            afternoondatetime = None

        if indexdate is None:
            indexdate = recordtime
        else:
            #日期发生变化，则记录打卡时间到表格中
            if indexdate.day != recordtime.day:

                print("morning time:%s"%name)
                print(moningdatetime)
                print(afternoondatetime)
                addnewattence(recordsheet, moningdatetime, afternoondatetime)

                #切换游标时间
                indexdate = recordtime
                moningdatetime = None
                afternoondatetime = None

        if moningdatetime is None or afternoondatetime is None:
            moningdatetime = recordtime
            afternoondatetime = recordtime
        else:
            #记录最早打卡和最晚打卡时间
            if moningdatetime.__ge__(recordtime):
                moningdatetime = recordtime
            if afternoondatetime.__le__(recordtime):
                afternoondatetime = recordtime
    recodewb.save()

def process_attendance(srcfile):
    try:
        workapp = xw.App(visible=True, add_book=False)
        #workapp.screen_updating = False
        core(workapp, srcfile)
    finally:
        #workapp.screen_updating = True
        workapp.quit()
    file_dir = os.path.dirname(srcfile)
    moveoneresultfile(file_dir, r'员工考勤时间表.xlsx')
