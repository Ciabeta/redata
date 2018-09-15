# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
#absolute index of files
jd15='15jdy.xlsx'
jd16='16jdy.xlsx'
jd17='17jdy.xlsx'


#open excel
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        #print str(e)
        print 'error'

#colnameindex：index of file  ，by_name：sheet name
def excel_table_byname(file, colnameindex, by_name):
    data = open_excel(file)
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
tables = excel_table_byname(jd16, 0, 'Sheet1')
for row in tables:
        print row[1]