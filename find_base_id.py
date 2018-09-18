# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
#absolute index of files
jd15 = '/root/.qqbot-tmp/plugins/15jdy.xlsx'
jd16 = '/root/.qqbot-tmp/plugins/16jdy.xlsx'
jd17 = '/root/.qqbot-tmp/plugins/17jdy.xlsx'

#open excel
def open_studata(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

#colnameindex：index of file  ，by_name：sheet name
def get_all_studata(file, colnameindex, by_name):
    data = open_studata(file)
    table = data.sheet_by_name(by_name)
    #colnames is a matrix of the specific row
    colnames = table.row_values(colnameindex)
    list =[]
    #read a table row by row
    for rownum in range(0, table.nrows):
        #get row
        row = table.row_values(rownum)
        if row:
            app = []
            #read a row col by col
            for i in range(len(colnames)):
               app.append(row[i])
            list.append(app)
    return list
def stu_infocheck(card_content):
        id_num = int(card_content)
        if(id_num <= 201500800001):
                return 0
	if(id_num <= 201600800001):
		return jd15
        elif(id_num <= 201700800001):
                return jd16
        elif(id_num <= 201800800001):
                return jd17
        else:
		return 0
def post_card_info(cardcontent):
	if(cardcontent.startswith('#')):
		cardcontent = cardcontent.strip('#')
	        file_select = stu_infocheck(cardcontent)    
	        tables = get_all_studata(file_select, 0, 'Sheet1')
	        for row in tables:
	            stuinfo_id = int(cardcontent)
	            if(row[2] == stuinfo_id):
	                return "他是%s的%s噢，请认识的同学通知一下吧！" % (row[0].encode("utf-8"), row[1].encode("utf-8"))
	else:
	    return 0;
def onQQMessage(bot, contact, member, content):
    if(contact.ctype == 'group'):
        if(contact.name == '找遍山威Beta'):
            return_stuinfo = post_card_info(content)
            if(return_stuinfo != 0):
                bot.SendTo(contact, return_stuinfo)
