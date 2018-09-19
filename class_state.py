# -*- coding: utf-8 -*-
from qqbot import qqbotsched
import  xdrlib ,sys
import xlrd
import datetime
class_jixie16 = '/root/data/class_jixie16.xlsx'
class_tongxin17 = '/root/data/class_tongxin17.xlsx'
def get_whichday():
    d = datetime.datetime.now()
    todayis = d.weekday()
    if(todayis == 1):
        return 1
    elif(todayis == 2):
        return 2
    elif(todayis == 3):
        return 3
    elif(todayis == 4):
        return 4
    elif(todayis == 5):
        return 5

def open_classdata(file_name):
    try:
        data = xlrd.open_workbook(file_name)
        return data
    except Exception,e:
        print str(e)
def get_nextclass(file_name, todayis, class_num):
    #only write for 2, 4, 6, 8 ...
    class_num = class_num*2 + 10;
    data = open_classdata(file_name)
    table = data.sheet_by_name('Sheet1')
    col = table.col_values(todayis)
    if(col[class_num] != ''):
        return "下节课是%s，上课地点是%s，请各位同学做好准备。" % (col[class_num].encode("utf-8"), col[class_num+1].encode("utf-8"))
    else:
    	print '这节课没课'
        return 0
@qqbotsched(day_of_week='0-4', hour='7', minute='20')
def class_state1(bot):
    gl = bot.List('group', '测试')
    if gl is not None:
        for group in gl:
            todayis = get_whichday()
            return_class = get_nextclass(class_jixie16, todayis, 1)
            if(return_class != 0):
                bot.SendTo(group, return_class)
    g2 = bot.List('group', '测试2')
    if g2 is not None:
        for group in g2:
            todayis = get_whichday()
            return_class = get_nextclass(class_tongxin17, todayis, 1)
            if(return_class != 0):
                bot.SendTo(group, return_class)
@qqbotsched(day_of_week='0-4', hour='9', minute='0')
def class_state2(bot):
    gl = bot.List('group', '测试')
    if gl is not None:
        for group in gl:
            todayis = get_whichday()
            return_class = get_nextclass(class_jixie16, todayis, 2)
            if(return_class != 0):
                bot.SendTo(group, return_class)
    g2 = bot.List('group', '测试2')
    if g2 is not None:
        for group in g2:
            todayis = get_whichday()
            return_class = get_nextclass(class_tongxin17, todayis, 2)
            if(return_class != 0):
                bot.SendTo(group, return_class)

@qqbotsched(day_of_week='0-4', hour='14', minute='0')
def class_state3(bot):
    gl = bot.List('group', '测试')
    if gl is not None:
        for group in gl:
            todayis = get_whichday()
            return_class = get_nextclass(class_jixie16, todayis, 3)
            if(return_class != 0):
                bot.SendTo(group, return_class)
    gl = bot.List('group', '测试2')
    if g2 is not None:
        for group in g2:
            todayis = get_whichday()
            return_class = get_nextclass(class_tongxin17, todayis, 3)
            if(return_class != 0):
                bot.SendTo(group, return_class)

@qqbotsched(day_of_week='0-4', hour='16', minute='0')
def class_state4(bot):
    gl = bot.List('group', '测试')
    if gl is not None:
        for group in gl:
            todayis = get_whichday()
            return_class = get_nextclass(class_jixie16, todayis, 4)
            if(return_class != 0):
                bot.SendTo(group, return_class)
    g2 = bot.List('group', '测试2')
    if g2 is not None:
        for group in g2:
            todayis = get_whichday()
            return_class = get_nextclass(class_tongxin17, todayis, 4)
            if(return_class != 0):
                bot.SendTo(group, return_class)

@qqbotsched(day_of_week='0-4', hour='18', minute='50')
def class_state5(bot):
    gl = bot.List('group', '测试')
    if gl is not None:
        for group in gl:
            todayis = get_whichday()
            return_class = get_nextclass(class_jixie16, todayis, 5)
            if(return_class != 0):
                bot.SendTo(group, return_class)
    g2 = bot.List('group', '测试2')
    if g2 is not None:
        for group in g2:
            todayis = get_whichday()
            return_class = get_nextclass(class_tongxin17, todayis, 5)
            if(return_class != 0):
                bot.SendTo(group, return_class)


