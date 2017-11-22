#encoding:utf-8
import sys
import xlrd
import xlwt
import os
global ROW
ROW = 1

def Write(table, rvalue, ip, port):
	global ROW
	table.write(ROW, 0, ip)
	table.write(ROW, 1, port)
	table.write(ROW, 2, rvalue[2])
	table.write(ROW, 3, rvalue[3])
	table.write(ROW, 4, rvalue[5])
	table.write(ROW, 5, rvalue[13])
	table.write(ROW, 6, rvalue[17])
	table.write(ROW, 7, rvalue[18])
	ROW += 1

def Deal(file):
	port = ''
	data = xlrd.open_workbook(file) #打开指定文件
	table = data.sheets()[1] #选定要提取工作表
	nrows = table.nrows #获取表总行数
	ip = file.split('.xls')[0]
	if nrows > 1:
		port = table.row_values(1)[0]
	for n in range(1, nrows):
		row_value = table.row_values(n)
		if row_value[0] != '' and row_value[0] != port:
			port = row_value[0]
		Write(result, row_value, ip, port)


#IP,端口，漏洞名称，风险等级，CVE编号，详细描述，解决办法
FILE = xlwt.Workbook() 
result = FILE.add_sheet('Sheet0', cell_overwrite_ok = True)
result.write(0, 0, 'IP')
result.write(0, 1, u'端口')
result.write(0, 2, u'服务')
result.write(0, 3, u'漏洞名称')
result.write(0, 4, u'风险等级')
result.write(0, 5, u'CVE编号')
result.write(0, 6, u'详细描述')
result.write(0, 7, u'解决办法')

List = os.listdir(os.getcwd())
for file in List:
	if(file.split('.')[-1] == 'xls'):
		Deal(file)

FILE.save('demo.xls')
