# -*- coding: utf-8 -*-
import sys
import xlrd
import os
from optparse import OptionParser
##########################################################
reload(sys)
sys.setdefaultencoding('utf-8')

##########################################################
def openExcel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)


##########################################################
def writeKeysValuesInToLocalizableFile(keys,values,targetFolder):
    if not os.path.exists(targetFolder):
        os.makedirs(targetFolder)

    fileName = targetFolder + 'Localizable.strings'
    os.system(r'touch %s' % fileName)
    fp = open(fileName,'wb+')

    keyValueList = []
    for indexRow in range(len(keys)):
        key = keys[indexRow]
        value = values[indexRow]
        keyValue = '"' + key + '"' + ' = ' + '"' + value + '"' + ';\n'
        keyValueList.append(keyValue)

    content = ''.join(keyValueList)
    fp.write(content)
    fp.close()

##########################################################
def importLocalizable(options):
    data = openExcel(options.filePath)
    print(options.filePath)
    if data :
        table = data.sheets()[0]
        colnames = table.row_values(0) #第一行数据
        colKeys = table.col_values(0) #第一列key数据
        print (colKeys)
        del colKeys[0]

        for indexCol in range(len(colnames)):
            if indexCol > 0:
                languageName = colnames[indexCol]
                values = table.col_values(indexCol)
                # print(values)
                del values[0]
                writeKeysValuesInToLocalizableFile(colKeys,values,os.getcwd()+"/iOSLocal/"+languageName+".proj/")
    else :
        print "can not open file"

##########################################################
#def importLocalizable1(options,options1):
#    targetFile = "localizableToExcel.xls"
#
#    data1 = openExcel(options.filePath)
#    data2 = openExcel("excel2")
#
#    if data1 :
#       if data2:
#        table1 = data1.sheets()[0]
#        colnames1 = table1.row_values(0) #第一行数据
#        colKeys1 = table1.col_values(0) #第一列key数据
#
#        table2 = data2.sheets()[0]
#        colnames2 = table2.row_values(0) #第一行数据
#        colKeys2 = table2.col_values(0) #第一列key数据
#        print (colKeys1)
#        del colKeys1[0]
#        colKeys3 = colKeys1 - colKeys2
#        print(colKeys3)
##        for indexRow in range(len(colKeys1))
##           if indexRow >0:
##              key1 = colKeys1[indexRow];
##              key2 = colKeys1[indexRow];
#
#
#
#        for indexCol in range(len(colnames)):
#            if indexCol > 0:
#                languageName = colnames[indexCol]
#                values = table.col_values(indexCol)
#                # print(values)
#                del values[0]
#                writeKeysValuesInToLocalizableFile(colKeys,values,os.getcwd()+"/iOSLocal/"+languageName+".proj/")
#    else :
#        print "can not open file"

##########################################################
#def exportToExcel():
#    directory = "iOSLocal"
#    targetFile = "localizableToExcel.xls"
#    if directory is not None:
#        index = 0
#        if targetFile is not None:
#            wb = xlwt.Workbook()
#            ws = wb.add_sheet('test',cell_overwrite_ok=True)
#
#            for parent, dirnames, filenames in os.walk(directory):
#                for dirname in dirnames:
#                    # Key 和 国家简码
#                    if index == 0:
#                        ws.write(0,0,"Key")
#                    # iOS 不同的本地化语言文件xx.proj/Localizable.strings xx 对应国际化的国家简码 eg:english -->en; zh-Hans; zh-Hant; vi; pt; fa;
#                    countryCode = dirname.split('.')[0]
#                    ws.write(0,index+1,countryCode)
#
#                    #Key 和value
#                    path = directory+'/'+dirname+'/Localizable.strings'
#                    (keys,values) = readKeysAndValuesFromeFilePath(path)
#                    # print(keys)
#                    # print('======================')
#                    # print(values)
#
#                    for x in range(len(keys)):
#                        key = keys[x]
#                        value = values[x]
#                        if (index == 0):
#                            ws.write(x+1, 0, key)
#                            ws.write(x+1, 1, value)
#                        else:
#                            ws.write(x+1, index + 1, value)
#                    index += 1
#
#            wb.save(targetFile)

##########################################################
##########################################################
def main():
    parser = OptionParser()
    parser.add_option("-f", "--filePath",
                      help="original.xls File Path.",
                      metavar="filePath")
    (options, args) = parser.parse_args()
    importLocalizable(options)

##########################################################
main()
