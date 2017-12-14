#/usr/bin/env python
#coding:utf-8
#update:20171209
#Author:kzaopa

import xlrd
import xlwt
import sys
import os
from xlutils.copy import copy
from xml.etree.ElementTree import parse

'''
最近遇到n多这样的情况，拿到一份资产表ip少则几百多则上千，需要扫描指定的端口再将端口状态与资产一一对应并新增列附在资产表中；
一开始是将ip提取出来用nmap/masscan扫描，结果保存为xml文档，从xml文档提取ip和端口状态，最后在与资产表中ip匹配并写入状态状态；
于是就有这个脚本的产生，把这几个步骤融合在一起，在这之前尝试过用socket访问端口并写入，感觉效率较低放弃了。

用法： cp为复制excel表并新增列，er为提取需要的列并新增
使用 cp 需要手动设置Xlutils_Copy_Read()函数中的 cols = t1.col_values(10) 和 ip = t1.cell(n, 10).value ；还有Write_cp()函数，扫描几个端口则新增几列
使用 er 需要手动设置Read()函数中的 cols = table.col_values(10) 和 if row_value[10]: ip = row_value[10] ； 还有Write_w()函数，需要提取哪些列则写入进入，扫描几个端口则新增几列
'''

reload(sys)
sys.setdefaultencoding('utf-8')

#调用masscan
def Masscan(ips, port):
    for ip in ips:
        conf = 'masscan -p%s %s -oX /tmp/masscan.xml --rate=300 --show close --echo >> /tmp/masscan.conf' % (port, ip.strip())
        os.system(conf)
    cmd = 'masscan -c /tmp/masscan.conf'
    os.system(cmd)

#调用nmap
def Nmap(ips, port):
    for ip in ips:
        with open('/tmp/nmap.txt', 'a') as fn:
            fn.write(ip.strip() + '\n')
    cmd = 'nmap -v -n -Pn -p%s -oX /tmp/nmap.xml -iL /tmp/nmap.txt' % port
    os.system(cmd)

#从nmap扫描结果xml文件中提取需要的数据
def Extract_xml_nmap(ip, *arg):
    values = []
    et = parse('/tmp/nmap.xml')
    root = et.findall('host')
    for child in root:
        if ip == child.find('address').attrib['addr']:
            for x in xrange(len(child.findall('ports/port'))):
                values.append([child.findall('ports/port')[x].attrib['portid'], child.findall('ports/port/state')[x].attrib['state']])
            return values

    #print child.find('address').attrib['addr'], child.findall('ports/port')[0].attrib['portid'], child.findall('ports/port/state')[0].attrib['state']

#从masscan扫描结果xml文件中提取需要的数据
def Extract_xml_masscan(ip, port):
    values = []
    z = 0
    count = len(port.split(','))
    et = parse('/tmp/masscan.xml')
    root = et.findall('host')
    for child in root:
        if z > count: break
        if ip == child.find('address').attrib['addr']:
            values.append([child.findall('ports/port')[0].attrib['portid'], child.findall('ports/port/state')[0].attrib['state']])
            z += 1
    return values


'''
利用Xlutils模块复制一个excel表并向其中追加新列
'''
#新增列，需要手动配置
def Write_cp(table, ROW, status= ''):
    #写入首行
    table.write(0, 19, u'端口号：21')
    table.write(0, 20, u'端口号：22')

    #按行写入数据；
    #加入判断是因为masscan扫描结果没有filtered状态

    if len(status) == 1:
        table.write(ROW, 19, status[0][1])
        table.write(ROW, 20, 'filtered')
    else:
        table.write(ROW, 19, status[0][1])
        table.write(ROW, 20, status[1][1])

#    if status[5][0] == '8443':
#        table.write(ROW, 24, status[5][1])

def Xlutils_Copy_Read(filename, scan, port, extract):

    data = xlrd.open_workbook(filename)
    t1 = data.sheets()[0]
    cols = t1.col_values(10)        #获取需要扫描的ip
    rows = t1.nrows     #获取总行数
    w = copy(data)
    wt = w.get_sheet(0)
    eval(scan)(cols, port)      #调用扫描函数
    for n in xrange(rows):
        ip = t1.cell(n, 10).value       #每行数据判断参数，一般为ip地址
        if not ip:      #ip为空则退出本次循环
            continue
        port_status = eval(extract)(ip, port)       #调用xml文档提取函数
        if not port_status:     #如果没有返回端口状态可能是端口状态全部为filtered，而masscan不能判断filtered而导致的
            continue
        Write_cp(wt, n, port_status)

    w.save('result.xls')


'''
利用xlrd模块读取excel表，并用xlwt模块生成一个全新的excel表
'''
def Write_w(table, ROW, value, status= ''):
    #global ROW
    #写入首行
    table.write(0, 0, u'所属业务系统')
    table.write(0, 1, u'所属安全域')
    table.write(0, 2, u'所属单位')
    table.write(0, 3, u'所属部门')
    table.write(0, 4, u'公网IP')
    table.write(0, 5, u'责任人')
    table.write(0, 6, u'责任人手机')
    table.write(0, 7, u'端口号：21')
    table.write(0, 8, u'端口号：22')
    
    #按行写入数据
    table.write(ROW, 0, value[2])
    table.write(ROW, 1, value[4])
    table.write(ROW, 2, value[6])
    table.write(ROW, 3, value[7])
    table.write(ROW, 4, value[10])
    table.write(ROW, 5, value[16])
    table.write(ROW, 6, value[17])
    if len(status) == 1:
        table.write(ROW, 7, status[0][1])
        table.write(ROW, 8, 'filtered')
    elif len(status) == 0:
        table.write(ROW, 7, 'filtered')
        table.write(ROW, 8, 'filtered')
    else:
        table.write(ROW, 7, status[0][1])
        table.write(ROW, 8, status[1][1])
    #ROW += 1

def Read(filename, scan, port, extract):
    file1 = xlwt.Workbook()
    newtable = file1.add_sheet('Sheet0', cell_overwrite_ok = True)
    data = xlrd.open_workbook(filename)     #读取excel表
    table = data.sheets()[0]
    nrows = table.nrows
    cols = table.col_values(10)        #获取需要扫描的ip
    eval(scan)(cols, port)      #调用扫描函数
    for n in xrange(nrows):
        row_value = table.row_values(n)
        ip = table.cell(n, 10).value
        if not ip:
            continue
        port_status = eval(extract)(ip, port)
        if not port_status:
            Write_w(newtable, n, row_value)
            continue
        Write_w(newtable, n, row_value, port_status)
    file1.save('result.xls')

#执行函数
def Action():
    d = {'cp': 'Xlutils_Copy_Read', 'er': 'Read', 'n': 'Nmap', 'm': 'Masscan', 'en': 'Extract_xml_nmap', 'em': 'Extract_xml_masscan'}
    try:
        mode = sys.argv[1]      #设置excel生成模式
        scanner = sys.argv[2]       #设置扫描器
        port = sys.argv[3]     #扫描端口
        extract = sys.argv[4]       #选择提取xml函数
        filename = sys.argv[5]      #需要处理源excel文件
    except:
        print 'Usage:   scan.py mode scanner port extract filename'
        print 'Example: scan.py er n 21,80 en demo.xls'
        print 'Options:'
        for k, v in d.items():
            print '          ' + k + ':' + v
        exit()

    if scanner == 'n' and extract == 'en':
        eval(d[mode])(filename, d[scanner], port, d[extract])
    elif scanner == 'm' and extract == 'em':
        eval(d[mode])(filename, d[scanner], port, d[extract])
    else:
        print 'option n only with en or m and em'
        exit()

Action()
