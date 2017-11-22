#/usr/bin/env python
#coding:utf-8
#data:20171117
#Author:kzaopa

import xlrd
import xlwt
import sys
import time
global ROW
ROW = 1

reload(sys)
sys.setdefaultencoding('utf-8')


def Write(table, value, status = ''):
    global ROW
    #写入首行
    sheet.write(0, 0, u'编号')
    sheet.write(0, 1, u'省份')
    sheet.write(0, 2, u'资产名称')
    sheet.write(0, 3, u'资产类别')
    sheet.write(0, 4, u'所属部门')
    sheet.write(0, 5, u'所属业务系统')
    sheet.write(0, 6, u'主识别IP地址')
    sheet.write(0, 7, u'是否开放443端口（第一次扫描时间：）')
    #按行写入数据
    table.write(ROW, 0, value[0])
    table.write(ROW, 1, value[1])
    table.write(ROW, 2, value[2])
    table.write(ROW, 3, value[3])
    table.write(ROW, 4, value[4])
    table.write(ROW, 5, value[5])
    table.write(ROW, 6, value[6])
    table.write(ROW, 7, status)
    ROW += 1


def Read(filename):
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    #newfile = time.time().split('.')[0] + '.xls'
    nrows = table.nrows
    for n in xrange(1, nrows):
        row_value = table.row_values(n)
        if row_value != '':
            ip = row_value[6]
            for i in open('123.txt').readlines():
                i = i.split('-')
                if ip == i[0]:
                    prot_status = i[3]
                    Write(sheet, row_value, prot_status)

file1 = xlwt.Workbook()
sheet = file1.add_sheet('Sheet0', cell_overwrite_ok = True)

filename = sys.argv[1]

Read(filename)
file1.save('result.xls')
