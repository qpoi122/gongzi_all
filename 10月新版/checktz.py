from __future__ import division
import xlrd
import os
import math
from xlwt import Workbook, Formula
import xlrd
import sys
import types
import copy  
import time  
import pandas

from datetime import datetime
from xlrd import xldate_as_tuple
def is_chinese(uchar): 
        """判断一个unicode是否是汉字"""
        if uchar >= u'/u4e00' and uchar<=u'/u9fa5':
                return True
        else:
                return False

                
def is_num(unum):
	try:
		unum+1
	except TypeError:
		return 0
	else:
		return 1

#不带颜色的读取
def filename(content):
	#打开文件
	global workbook,file_excel
	file_excel=str(content)
	file=(file_excel+'.xls').decode('utf-8')#文件名及中文合理性
	if not os.path.exists(file):#判断文件是否存在
		file=(file_excel+'.xlsx').decode('utf-8')
		if not os.path.exists(file):
			print "文件不存在"
	workbook = xlrd.open_workbook(file)
	print 'suicce'


def readexcel(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		
		table = workbook.sheets()[name]
		ttype=table.name
		nrows=table.nrows
		for n in range(nrows):			
		#获取每行内容
			a=table.row_values(n)
			mid=[]
			mid.append(ttype)
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数 
				mid.append(a[i])
			money.append(mid)

	# print money,'montytyyt'
def readexcel2(content):
	filename(content)

		#获取所有的sheet
	Sheetname=workbook.sheet_names()
	# print "文件",file_excel,"共有",len(Sheetname),"个sheet："
	for name in range(len(Sheetname)):
		
		table = workbook.sheets()[name]
		ttype=table.name
		nrows=table.nrows
		midd=[]
		midd.append(ttype)
		for n in range(nrows):			
		#获取每行内容
			a=table.row_values(n)
			mid=[]
	
			for i in range(len(a)):	
						
				if is_chinese(a[i]):
					a[i].encode('utf-8' )
				elif is_num(a[i])==1:
					if math.modf(a[i])[0]==0 or a[i]==0:#获取数字的整数和小数
						a[i]=int(a[i])#将浮点数化成整数 
				mid.append(a[i])
			midd.append(mid)
		taizhang.append(midd)
	print taizhang,'taizhang'

# def addsi(zhangdan,money):

if __name__ == "__main__":
	global zhangdan,money,gongzi,errortype,renming,needi,taizhang
	zhangdan=[]
	money=[]
	gongzi=[]
	errortype=[]
	renming=[]
	needi=[]
	taizhang=[]
	readexcel1('newmy')
	readexcel('12345')
	readexcel2('taizhang')
	# readexcel('2345')
	# readexcel2('tztztztztztztz')