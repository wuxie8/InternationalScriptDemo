# -*- coding: utf-8 -*-
import sys
import xdrlib
import xlrd
import os
import shutil
import xlwt

##########################################################
reload(sys)
sys.setdefaultencoding('utf-8')

##########################################################
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
#        print(data)
        return data
    except Exception,e:
        print str(e)

##########################################################
def main(argv,argv1):
    data = open_excel(argv)
    targetFile = "localizableToExcel117.xls"
    index = 0
    indexkey = 0

    data1 = open_excel(argv1)

    if data:
#    if data1:

        table = data.sheets()[0]
        colnames = table.row_values(0) #第一行数据
        

        colKeys = table.col_values(0) #第一列key数据
        
        colKeys10 = table.col_values(1) #第二列key数据
        
        colKeys20 = table.col_values(2) #第3列key数据


        table1 = data1.sheets()[0]
        colnames1 = table1.row_values(0) #第一行数据
        
        colKeys1 = table1.col_values(0) #第一列key数据
        
        colKeys11 = table1.col_values(1) #第二列key数据
        
        colKeys21 = table1.col_values(2) #第3列key数据

        colKeys2 = set(colKeys) - set(colKeys1)
        colKeys12 = set(colKeys10) - set(colKeys11)
        
        colKeys3 = list(colKeys2)
        colKeys13 = list(colKeys12)

#        print(table.col_values)
#        print("***")
#
#        print(colKeys20[3])
#        print("===")
#
        print(colKeys[10])
        print(colKeys10[10])
        print(colKeys20[10])

#        print(colKeys3[0])
#        print(colKeys13[0])
#        print(colKeys3[0])
#        print(colKeys13[0])
        if targetFile is not None:
            wb = xlwt.Workbook()
            ws = wb.add_sheet('test',cell_overwrite_ok=True)
            for x in range(len(colKeys)):
                key = colKeys[x]
                if len(key)>0:
#                    for y in range(len(colKeys1))
#                     indexkey1 = colKeys1.index(key)
#                print(index)
#                print(key)
#                print("------")
#                print(colKeys10[index])
                    if key in colKeys1:
                        indexkey1 = colKeys1.index(key)

                        value = colKeys1[indexkey1]
                        value1 = colKeys21[indexkey1]
                        value2 = colKeys11[indexkey1]

                        print(key)
                        print(value1)
                        print(value2)

                        if (index == 0):
                            value3 = str(value1)
                            if len(value3)>0:
                                ws.write(indexkey, 0, value)
                                ws.write(indexkey, 1, value2)
                                ws.write(indexkey, 2, value1)

                            else:
                                ws.write(indexkey, 0, value)
                                ws.write(indexkey, 1, value2)

                                ws.write(indexkey, 2, value1)
                        else:
                                ws.write(indexkey, index + 1, value)
                                index += 1
                    else:
                        value1 = colKeys20[x]
                        value2 = colKeys10[x]

                        print(key)
                        print(value1)
                        print(value2)

                        if (index == 0):
                            value3 = str(value1)

                            if len(value3)>0:
                                ws.write(indexkey, 0, key)
                                ws.write(indexkey, 1, value2)
                                ws.write(indexkey, 2, value1)
                            else:
                                ws.write(indexkey, 0, key)
                                ws.write(indexkey, 1, value2)
                                ws.write(indexkey, 2, value1)
                                #                        ws.write(indexkey, 2, value1)
                        else:
                                ws.write(indexkey, index + 1, value)
                                index += 1
                indexkey +=1
        wb.save(targetFile)

#        colValues_zh_CN = table.col_values(1) #简体中文数据
#        colValues_English = table.col_values(2)#英文数据
#        nrows = len(colKeys) #总行数
#        ncols = len(colnames) #总列数
        
#        languageList = []
#        for indexCol in range(1,ncols):
#            list1 = []
#            colValues = table.col_values(indexCol)
#            for indexRow in range(1,nrows):
#
#                value = colValues[indexRow]
#                if (len(str(value))==0):
#                    value = colValues_English[indexRow]
#
##                keyValue = '"' + colKeys[indexRow] + '"' + ' = ' + '"' + value + '"' + ';\n'
#                keyValue = '"' + str(colKeys[indexRow]) + '"' + ' = ' + '"' + str(value) + '"' + ';\n'
#
#
#                list1.append(keyValue)
#            languageList.append(''.join(list1))
#
#
#        for index in range(len(languageList)):
##            print languageList[index]
#            fileName = str(index) + 'Localizable.strings'
#            os.system(r'touch %s' % fileName)
#
#            fp = open(fileName,'wb+')
#            fp.write(languageList[index])
#            fp.close()
#    else :
#                print "can not open file"

if __name__=="__main__":
  

    main(sys.argv[1],sys.argv[2])
