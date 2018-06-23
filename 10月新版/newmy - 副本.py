# -*- coding: UTF-8 -*-

'''
__author__="zf"
__mtime__ = '2016/11/8/21/38'
__des__: 简单的读取文件
__lastchange__:'2016/11/16'
'''
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
    if str(uchar) >= '/u4e00' and str(uchar) <= '/u9fa5':
        return True
    else:
        return False


def is_num(unum):
    try:
        unum + 1
    except TypeError:
        return 0
    else:
        return 1


# 不带颜色的读取
def filename(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    file = (file_excel + '.xls')  # 文件名及中文合理性
    if not os.path.exists(file):  # 判断文件是否存在
        file = (file_excel + '.xlsx')
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    print('suicce')


# 读取帐目，生成的是标题和内容一对一的
def readexcel(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):
        table = workbook.sheets()[name]
        nrows = table.nrows
        title = table.row_values(1)
        print(title, 'titititiitittiitititiitit')
        for x in range(len(title)):

            # print title[x],'sadjioajdoi'
            if u'.' in title[x]:

                gongxuleixing = title[x].split(u'.')[0]
                for y in renming:
                    # print y
                    if int(gongxuleixing) == y[0]:
                        zuhemingcheng = title[x].split(u'.')[1]
                        print(zuhemingcheng, 'ssssssss')
                        if u'/' in zuhemingcheng:
                            print('123213')
                            splitxingming = zuhemingcheng.split(u'/')
                            for z in splitxingming:
                                if z not in y:
                                    y.append(z)
                        else:
                            y.append(zuhemingcheng)

        for x in renming:
            del x[0]

        print(renming, 'rrenminh@@#@#')

        # for x in range (len(title)):
        # 	ren=[]
        # 	flag=0

        # 	if u'/' in title[x]:
        # 		ren=title[x].split(u'/')
        # 	elif u'和' in title[x]:
        # 		ren=title[x].split(u'和')
        # 	else:
        # 		for z in pd:
        # 			if z ==title[x]:
        # 				flag=1
        # 				break
        # 		if flag==0:
        # 			ren=title[x]
        # 		else:
        # 			ren=[]

        # 	mid=[]
        # 	# mid.append(x)
        # 	if ren!=[]:
        # 		for y in ren:
        # 			mid.append(y)
        # 		renming.append(mid)
        # print renming,'renrenrnenrnernenr'

        # 获取每行的信息加上抬头
        # print title,'title'
        for n in range(2, nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []
            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')

                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数
                try:
                    a[i] = a[i].strip()
                except:
                    pass

                mid.append(title[i])
                mid.append(a[i])

            zhangdan.append(mid)
        # print renming,'renrenrnenrenrss'
    # print renming,'sdsdsdsd'
    # print zhangdan,'zhangdan'
    # diyi=0
    # for x in range(len(zhangdan)):
    # 	diyi=diyi+zhangdan[x][u'第一']
    # print diyi,'yi'

    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')
    for i in range(len(zhangdan)):
        for j in range(len(zhangdan[i])):
            if is_chinese(zhangdan[i][j]):
                zhangdan[i][j].encode('utf-8')
            # elif not zhangdan[i] and zhangdan[i]!=0:
            # 	print "空值",
            elif is_num(zhangdan[i][j]) == 1:
                if math.modf(zhangdan[i][j])[0] == 0 or zhangdan[i][j] == 0:  # 获取数字的整数和小数
                    zhangdan[i][j] = int(zhangdan[i][j])  # 将浮点数化成整数
            sheet1.write(i, j, zhangdan[i][j])
    book.save('allmesg.xls')  # 存储excel
    book = xlrd.open_workbook('allmesg.xls')


def readexcel1(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        for n in range(nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []
            mid.append(ttype)
            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')
                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数
                mid.append(a[i])
            money.append(mid)





def readexcel2(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []

            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')
                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数
                mid.append(a[i])
            midd.append(mid)
        taizhang.append(midd)
    print(taizhang, 'taizhang')


# def addsi(zhangdan,money):

# def new_
def addname():
    # n=1
    # while True:
    #     yield n
    #     n+=2yield n
    for x in renming:
        aa = dict.fromkeys(x, [])
        yield aa


# def new_
def addnewname():
    # n=1
    # while True:
    #     yield n
    #     n+=2yield n
    for x in smulu:
        aa = dict.fromkeys(x, [])
        yield aa


def readnodanzi(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []
            if a[-1] != u'o':
                for i in range(len(a)):

                    if is_chinese(a[i]):
                        a[i].encode('utf-8')
                    elif is_num(a[i]) == 1:
                        if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                            a[i] = int(a[i])  # 将浮点数化成整数
                    mid.append(a[i])
                midd.append(mid)
        nodanzi.append(midd)
    print(nodanzi, 'nodanzinodanzinodanzinodanzinodanzinodanzi')


def readnotaizhang(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []
            if a[-1] == u'x':
                for i in range(len(a)):

                    if is_chinese(a[i]):
                        a[i].encode('utf-8')
                    elif is_num(a[i]) == 1:
                        if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                            a[i] = int(a[i])  # 将浮点数化成整数
                    mid.append(a[i])
                midd.append(mid)
        notaizhang.append(midd)


def readexcel2(content):
    filename(content)

    # 获取所有的sheet
    Sheetname = workbook.sheet_names()
    # print "文件",file_excel,"共有",len(Sheetname),"个sheet："
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):
            # 获取每行内容
            dierhang = u''
            a = table.row_values(n)
            mid = []

            for i in range(len(a)):

                if is_chinese(a[i]):

                    a[i].encode('utf-8')
                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数

                if type(a[i]) == type(u''):
                    try:
                        a[i] = int(a[i])
                    except:
                        pass

                if i == 1:
                    dierhang = a[i]
                elif i == 2:
                    mid.append(a[i])
                    mid.append(dierhang)
                else:
                    mid.append(a[i])
            midd.append(mid)
        taizhang.append(midd)


if __name__ == "__main__":
    global zhangdan, money, gongzi, errortype, renming, needi, taizhang, nodanzi, notaizhang
    zhangdan = []
    money = []
    gongzi = []
    errortype = []
    # renming=[]
    # renming = [[1, 1], [2, 2], [3, 3], [4, 4]]
    renming = [[1], [2], [3], [4]]
    needi = []
    taizhang = []
    nodanzi = []
    notaizhang = []
    readexcel1('newmy')
    readexcel('12345')
    readexcel2('taizhang')
    # readexcel('2345')
    # readexcel2('tztztztztztztz')

    add_name = addname()
    for o in add_name:
        needi.append(o)

    # 根据获取到的人名,再次遍历账单,将碰到的信息分配给对应的人下面,并且去比对money表加上钱

    for x in zhangdan:
        for z in range(len(x)):
            # 提前获取各个信息
            pd = [u'客户编号', u'总成型号', u'数量', u'日期', u'塑料袋', u'小标贴', u'小内盒']

            # zcbh=[]
            # khbh=[]
            # print x,'xxxxxxxxxxxxxxxxxxxxx'
            if x[z] == u'总成型号':
                # print 'qioeuoqidoaih'
                zcbh = x[z + 1]
            elif x[z] == u'客户编号':
                khbh = x[z + 1]
            elif x[z] == u'塑料袋':
                sld = x[z + 1]
            elif x[z] == u'小标贴':
                xbt = x[z + 1]
            elif x[z] == u'小内盒':
                xnh = x[z + 1]

        for z in range(len(x)):
            pipei_num = []
            if x[z] not in pd and z % 2 == 0:
                try:
                    namewithno = x[z].split(u'.')[1]
                    if u'/' in namewithno:
                        pipei_num = namewithno.split(u'/')
                    elif u'和' in namewithno:
                        pipei_num = namewithno.split(u'和')
                    else:
                        pipei_num.append(namewithno)
                except:
                    pass

            for m in range(len(pipei_num)):
                # 将a中匹配到的键，值取出，然后增加值，然后放回去
                for h in range(len(needi)):
                    if pipei_num[m] in needi[h] and x[z + 1] != u'' and x[z + 1] != u' ' and type(
                            x[z + 1]) != type(u''):
                        zhongji = []  # zong
                        mid = []  # xiao
                        zhongji = copy.deepcopy(needi[h].get(pipei_num[m]))

                        # 加上匹配到的项的日期
                        flgggggg = 0
                        a = z - 1
                        for p in range(0, 2):
                            # print x[a],'wdwdwdwdwdwd'
                            if type(x[a]) == type(1):
                                mid.append(x[a])
                                flgggggg = 1
                                break
                            else:
                                a = a - 1
                            # print '11111'
                        if flgggggg == 0:
                            mid.append(u'no tiqi')

                        # mid.append(x[1])

                        mid.append(zcbh)
                        mid.append(khbh)
                        if type(x[z + 1]) != type(1):
                            print(z, x[z + 1], type(x[z + 1]), type(len(pipei_num)))

                        # 这里可能应为汉字报错

                        # #划分数量
                        # 						shul=x[z+1]/len(pipei_num)
                        # 						mid.append(shul)

                        # 						zongmoney=0
                        # 						finded=0
                        # 						for j in range(len(money)):
                        # 							# if h!=0:
                        # 							# 	print int(money[j][0]),h,type(int(money[j][0])),type(h)
                        # 								if int(money[j][0])==h+1 and money[j][1]==x[3] and int(money[j][0])!=4:
                        # 									# print 'ssssssssssssssssss',money[j][2],shul,type(money[j][2])
                        # 									try:
                        # 										zongmoney=money[j][2]*shul
                        # 									except:
                        # 										pass
                        # 									# print zongmoney
                        # 									break

                        # 						# if h+1==4 and finded==0:
                        # 						# 	for j in range(len(money)):
                        # 						# 		flag=0
                        # 						# 		if zhan
                        mid.append(len(pipei_num))
                        # 划分单价
                        shul = x[z + 1]
                        mid.append(shul)

                        zongmoney = 0
                        finded = 0
                        for j in range(len(money)):
                            # if h!=0:
                            # 	print int(money[j][0]),h,type(int(money[j][0])),type(h)

                            # 将除了第四道的工序都匹配金钱,没有匹配到的钱为0
                            if int(money[j][0]) == h + 1 and money[j][1] == x[3] and int(money[j][0]) != 4:
                                # print 'ssssssssssssssssss',money[j][2],shul,type(money[j][2])
                                try:
                                    zongmoney = money[j][2] * shul / len(pipei_num)
                                except:
                                    pass
                                # print zongmoney
                                break

                        mid.append(zongmoney)
                        # 如果是第四种加上√ 信息
                        if h == 3:
                            # mid.append(x[23])
                            # mid.append(x[25])
                            # mid.append(x[27])

                            mid.append(sld)
                            mid.append(xbt)
                            mid.append(xnh)
                        zhongji.append(mid)
                        needi[h][pipei_num[m]] = zhongji
                    # print zhongji,needi[h]

    smulu = []
    for x in needi:
        dlist = list(x.keys())
        smulu = smulu + dlist

    # 核对一下账目和台账是否匹配
    lackzm = []

    for x in smulu:
        aa = {x: []}
        lackzm.append(aa)

    name4 = renming[-1]

    lacktj = []
    book1 = Workbook()

    for x in smulu:

        zmsj = []
        for y in needi:

            if x in y:

                zmsj = copy.deepcopy(y.get(x))
                for z in zmsj:
                    # print z,'zzzzzzzzz'
                    del z[3]
                    z.pop()
                    if x in name4:
                        del z[4:]
                break
        # print x,zmsj,'zmzmzmzmsjsjsjs'

        tj = []
        for z in range(len(taizhang)):
            # print taizhang[z][0],'what the fuck'
            if x in taizhang[z][0]:
                tj = copy.deepcopy(taizhang[z])
                break
        # print x,tj,'tjtjtjtjtjtjtj'

        # time.sleep(5)
        if tj != []:
            for m in range(len(zmsj)):
                flag = 0
                for n in range(len(tj)):
                    if zmsj[m][1] == tj[n][1] and zmsj[m][2] == tj[n][2] and zmsj[m][3] == tj[n][3] and tj[n][-1] != u'o':
                        try:
                            if 0 <= abs(zmsj[m][0] - tj[n][0]) <= 5 or 365 <= abs(zmsj[m][0] - tj[n][0]) <= 371:
                                # print type(zmsj[m][0]),type(tj[n][0]),tj[n][0],zmsj[m][0],'47897949',tj[n][0]-zmsj[m][0]
                                flag = 1
                                zmsj[m].append(u'o')
                                tj[n].append(u'o')
                                break
                        except:
                            pass

                if flag == 0:
                    zmsj[m].append(u'x')

        for mm in lackzm:
            if x in mm:
                dd = mm.get(x)
                for k in zmsj:
                    dd.append(k)

        # titles=u'zm'+x
        titles = x
        sheet1 = book1.add_sheet(titles)
        for i in range(1, len(tj)):
            for j in range(len(tj[i])):
                if is_chinese(tj[i][j]):

                    tj[i][j].encode('utf-8')

                # elif not tj[i] and tj[i]!=0:
                # 	print "空值",
                elif is_num(tj[i][j]) == 1:
                    if math.modf(tj[i][j])[0] == 0 or tj[i][j] == 0:  # 获取数字的整数和小数
                        tj[i][j] = int(tj[i][j])  # 将浮点数化成整数

                # print tj[i][j]
                if j == 0:
                    try:
                        cccc = pandas.to_datetime(tj[i][j] - 25569, unit='d')
                        tj[i][j] = str(pandas.Period(cccc, freq='D'))
                    except:
                        pass
                # print i , j,tj[i][j]
                sheet1.write(i - 1, j, tj[i][j])
    # print lackzm,'sdhushdiushidhsihdisuhdiushdiu'
    book1.save('no danzi.xls')  # 存储excel
    book1 = xlrd.open_workbook('no danzi.xls')

    book2 = Workbook()
    for x in smulu:
        sheet1 = book2.add_sheet(x)
        for y in lackzm:
            if x in y :
                sss = copy.deepcopy(y.get(x))
                for i in range(len(sss)):
                    for j in range(len(sss[i])):
                        # print 'xxxxx'
                        if is_chinese(sss[i][j]):

                            sss[i][j].encode('utf-8')

                        # elif not sss[i] and sss[i]!=0:
                        # 	print "空值",
                        elif is_num(sss[i][j]) == 1:
                            if math.modf(sss[i][j])[0] == 0 or sss[i][j] == 0:  # 获取数字的整数和小数
                                sss[i][j] = int(sss[i][j])  # 将浮点数化成整数

                        # print sss[i][j]
                        if j == 0:
                            try:
                                cccc = pandas.to_datetime(sss[i][j] - 25569, unit='d')
                                sss[i][j] = str(pandas.Period(cccc, freq='D'))
                            except:
                                pass
                        sheet1.write(i, j, sss[i][j])
    book2.save('no taizhang.xls')  # 存储excel
    book2 = xlrd.open_workbook('no taizhang.xls')

    # 将核对好的，不匹配的部分进行进一步的模糊匹配
    readnodanzi('no danzi')
    readnotaizhang('no taizhang')
    pijiaoout = []

    # for j in range(len(nodanzi)):
    #     mid3 = []
    #     mid3.append(nodanzi[j][0])
    #     for x in range(1, len(nodanzi[j])):
    #
    #         for y in range(1, len(notaizhang[j])):
    #             flag = 0
    #             mid = []
    #             mid.extend(nodanzi[j][x])
    #             mid.append(u'')
    #             mid.append(u'')
    #             if u'ok' not in mid:
    #                 mid.append(u'')
    #             mid.append(u'')
    #             mid.extend(notaizhang[j][y])
    #
    #             mid2 = []
    #             for z in mid:
    #                 mid2.append(u'')
    #
    #             # print nodanzi[j][x],notaizhang[j][y],'xyxyxyyxyyxyxyyxxyy'
    #
    #             if nodanzi[j][x][0] != notaizhang[j][y][0]:
    #                 flag += 10
    #                 mid2[0] = mid2[9] = u'xx'
    #             if nodanzi[j][x][1] != notaizhang[j][y][1]:
    #                 flag += 10
    #                 strnodanzi = str(nodanzi[j][x][1])
    #                 strnotaizhang = str(notaizhang[j][y][1])
    #                 if abs(len(strnodanzi) - len(strnotaizhang)) > 0:
    #                     flag = flag + 2 * abs(len(strnodanzi) - len(strnotaizhang))
    #
    #                 for z in range(min(len(strnodanzi), len(strnotaizhang))):
    #                     if strnodanzi[z] != strnotaizhang[z]:
    #                         flag += 2
    #                 mid2[1] = mid2[10] = u'xx'
    #             if nodanzi[j][x][2] != notaizhang[j][y][2]:
    #                 flag += 10
    #                 mid2[2] = mid2[11] = u'xx'
    #             if nodanzi[j][x][3] != notaizhang[j][y][3]:
    #                 flag += 10
    #                 mid2[3] = mid2[12] = u'xx'
    #             if flag <= 14:
    #                 mid3.append(mid)
    #                 mid3.append(mid2)
    #                 nodanzi[j][x].append(u'ok')
    #                 notaizhang[j][y].append(u'ok')
    #
    #     pijiaoout.append(mid3)
    for j in range(len(nodanzi)):
        mid3 = []
        mid3.append(nodanzi[j][0])
        for x in range(1, len(nodanzi[j])):
            for y in range(1, len(notaizhang[j])):
                flag = 0
                if nodanzi[j][x][0] != notaizhang[j][y][0]:
                    flag += 10
                if nodanzi[j][x][1] != notaizhang[j][y][1]:
                    flag += 10
                if nodanzi[j][x][2] != notaizhang[j][y][2]:
                    flag += 10
                if nodanzi[j][x][3] != notaizhang[j][y][3]:
                    flag += 15
                if flag <= 14:
                    nodanzi[j][x][-1] = u'ok'
                    notaizhang[j][y].append(u'ok')
                elif flag == 15:
                    mid = []
                    mid.extend(nodanzi[j][x])
                    mid.append(u'')
                    mid.append(u'')
                    if u'ok' not in mid:
                        mid.append(u'')
                    mid.append(u'')
                    mid.extend(notaizhang[j][y])
                    mid2 = []
                    for z in mid:
                        mid2.append(u'')
                    mid3.append(mid)
                    mid3.append(mid2)
                    nodanzi[j][x][-1] = u'ok'
                    notaizhang[j][y].append(u'ok')
        pijiaoout.append(mid3)
    print(pijiaoout, 'pijianajiajaiajiajaiaji')

    book5 = Workbook()

    for z in range(len(pijiaoout)):
        sheet1 = book5.add_sheet(pijiaoout[z][0])
        for i in range(1, len(pijiaoout[z])):
            for j in range(len(pijiaoout[z][i])):
                if is_chinese(pijiaoout[z][i][j]):
                    pijiaoout[z][i][j].encode('utf-8')
                # elif not pijiaoout[z][i] and pijiaoout[z][i]!=0:
                # 	print "空值",
                elif is_num(pijiaoout[z][i][j]) == 1:
                    if math.modf(pijiaoout[z][i][j])[0] == 0 or pijiaoout[z][i][j] == 0:  # 获取数字的整数和小数
                        pijiaoout[z][i][j] = int(pijiaoout[z][i][j])  # 将浮点数化成整数
                sheet1.write(i, j, pijiaoout[z][i][j])
    book5.save('mohu.xls')  # 存储excel
    book5 = xlrd.open_workbook('mohu.xls')

    # #模糊匹配失败的，查看前后是否有同型号，同客户，同数量的也算漏记
    # qianhoupipie=[]
    nodanzicopy = copy.deepcopy(nodanzi)

    for x in range(len(nodanzi)):
        for y in range(len(nodanzicopy)):
            for z in range(1, len(nodanzi[x])):
                for i in range(1, len(nodanzicopy[y])):
                    if len(nodanzi[x]) > 1 and len(nodanzicopy[y]) > 1:
                        if nodanzi[x][z][1:] == nodanzicopy[y][i][1:] and nodanzi[x][0] != nodanzicopy[y][0]:
                            if (not (nodanzi[x][0] in renming[0] and nodanzicopy[y][0] in renming[0])):
                                if  nodanzi[x][z][-1] !='ok' and nodanzicopy[x][z][-1] !='ok':
                                    nodanzi[x][z][-1] = u'find'

    # 模糊匹配都失败的单独拿出来
    book6 = Workbook()

    for z in range(len(nodanzi)):
        sheet1 = book6.add_sheet(nodanzi[z][0])
        line = -1
        for i in range(1, len(nodanzi[z])):
            if nodanzi[z][i][-1] != u'ok':
                line = line + 1
                for j in range(len(nodanzi[z][i])):
                    if is_chinese(nodanzi[z][i][j]):
                        nodanzi[z][i][j].encode('utf-8')
                    # elif not nodanzi[z][i] and nodanzi[z][i]!=0:
                    # 	print "空值",
                    elif is_num(nodanzi[z][i][j]) == 1:
                        if math.modf(nodanzi[z][i][j])[0] == 0 or nodanzi[z][i][j] == 0:  # 获取数字的整数和小数
                            nodanzi[z][i][j] = int(nodanzi[z][i][j])  # 将浮点数化成整数

                    # print line,'lineieni n'
                    sheet1.write(line, j, nodanzi[z][i][j])




        line = -1
        for i in range(1, len(notaizhang[z])):
            if notaizhang[z][i][-1] != u'ok':
                line = line + 1
                for j in range(len(notaizhang[z][i])):
                    if is_chinese(notaizhang[z][i][j]):
                        notaizhang[z][i][j].encode('utf-8')
                    # elif not notaizhang[z][i] and notaizhang[z][i]!=0:
                    # 	print "空值",
                    elif is_num(notaizhang[z][i][j]) == 1:
                        if math.modf(notaizhang[z][i][j])[0] == 0 or notaizhang[z][i][j] == 0:  # 获取数字的整数和小数
                            notaizhang[z][i][j] = int(notaizhang[z][i][j])  # 将浮点数化成整数

                    sheet1.write(line, j + 6, notaizhang[z][i][j])
    book6.save('lackmohu.xls')  # 存储excel
    book6 = xlrd.open_workbook('lackmohu.xls')

    # 对第四道进行单独处理,先匹配名字一样的,没有再去按老方法处理

    for x in smulu:
        if x in needi[3]:
            # print '111'
            for y in needi[3].get(x):

                # print '2222'
                for z in range(len(money)):
                    # print '3333'
                    if money[z][0] == u'4' and y[1] == money[z][1]:
                        # print y,y[6],'y6'
                        duigeshu = len(y[7])
                        jiage = ((duigeshu - 1) * 0.005 + money[z][2]) / y[3]
                        y[5] = jiage * y[4]
                    # print jiage,y[3],y[1],y[4],'1111111111111'

    for x in smulu:
        if x in needi[3]:
            # print '111'
            for y in needi[3].get(x):
                flag = 0
                for z in range(len(money)):
                    # print '3333'
                    if money[z][0] == u'5' and y[5] == 0:
                        # print '4444444444'
                        # 内盒一样并且带-的
                        # print len(y),len(money[z]),(y),money[z]
                        # print y
                        if y[8] == money[z][3] and u'-' == money[z][1] and type(y[1]) == type(u'a'):
                            # print '11111111111111111111111111111'
                            # if (type(zhangdan[x][3])==type(u'a') and len(zhangdan[x][3])>6):
                            jiage = money[z][4] / y[3]
                            y[5] = jiage * y[4]
                            flag = 1
                            # print y[4],'y3yy33yy3y3y3y3'
                            break
                        # 所有带几盒几个的,但是不带-的
                        elif u'个' in y[7] or u'只' in y[7]:
                            if (type(y[1]) == type(u'a') and u'-' not in y[1]) or type(y[1]) == type(1):

                                if y[6] == money[z][1] and y[8] == money[z][3]:
                                    # print '1111111111111111111111111111'
                                    if y[7] == '' and y[7] == money[z][2]:
                                        # print '222222222222222222222222222222222222222'
                                        jiage = money[z][4] / y[3]
                                        y[5] = jiage * y[4]
                                        flag = 1
                                        break
                                    elif y[7] != '' and money[z][2] == u'√':
                                        geshu = len(y[7])
                                        jiage = ((geshu - 1) * 0.005 + money[z][4]) / y[3]
                                        y[5] = jiage * y[4]
                                        flag = 1
                                        break
                        # 带排/数的
                        elif (type(y[1]) == type(u'a') and u'-' not in y[1] and (y[8] == u'排' or y[8] == u'数')) or (
                                type(y[1]) == type(1) and (y[8] == u'排' or y[8] == u'数')):
                            # print type(y[1])==type(u'a'),y[1],'sdsadsadsdasd'
                            s = [u'A', u'B', u'C', u'D', u'E', u'F', u'G', u'H']
                            youzimu = 0
                            lenxinhao = 0
                            if type(y[1]) == type(u'a'):
                                for p in range(len(s)):
                                    if s[p] in y[1]:
                                        youzimu = 1
                                        break

                            if youzimu == 1:
                                lenxinhao = len(y[1]) - 1
                            else:
                                try:
                                    lenxinhao = len(y[1])
                                except:
                                    if y[1] > 10000 and y[1] < 30000:
                                        lenxinhao = 5
                                    else:
                                        lenxinhao = 4

                            if lenxinhao == money[z][1] and y[8] == money[z][3]:
                                jiage = money[z][4] / y[3]
                                y[5] = jiage * y[4]
                                flag = 1
                                break

                        # 剩余的正常型号
                        elif y[6] == money[z][1] and y[8] == money[z][3]:
                            # print y[1],'SUSUSUSUSSU'
                            if y[7] == '' and y[7] == money[z][2]:
                                jiage = money[z][4] / y[3]
                                y[5] = jiage * y[4]
                                flag = 1
                                break
                            elif y[7] != '' and money[z][2] == u'√':
                                geshu = len(y[7])
                                jiage = ((geshu - 1) * 0.005 + money[z][4]) / y[3]
                                y[5] = jiage * y[4]
                                flag = 1
                                # print y[1],jiage,geshu,money[z][4],'2323232323'
                                break

                # else:
                # 	flag1=1
                # 	mid=[]
                # 	mid.append(zhangdan[x][5])
                # 	mid.append(zhangdan[x][3])
                # 	mid.append(u'4')
                # 	errortype.append(mid)
                # 	break

    # print needi,'sssssssssssss'

    # 补差价
    tijia = []
    for x in range(len(money)):
        if money[x][0] == u'6':
            tijia.append(money[x])

    for x in range(len(tijia)):
        # print renming[tijia[x][1]-1],'11111111111111111bbbbbbbbbbb'
        for y in renming[tijia[x][1] - 1]:
            for z in needi[tijia[x][1] - 1].get(y):
                # print z[4],tijia[x][1],'??????????????'
                if u'' in tijia[x]:
                    flag = 0
                    for j in range(len(tijia[x])):
                        if tijia[x][j] == tijia[x][1]:
                            if z[4] <= tijia[x][j + 1]:
                                newjiage = z[5] * (1 + tijia[x][j + 2])
                                # print newjiage
                                flag = 1
                                z.append(newjiage)
                                break
                    if flag == 0:
                        z.append(z[5])

                elif z[4] <= tijia[x][2]:
                    # print 'ddddddddddddddddddd'
                    newjiage = z[5] * (1 + tijia[x][3])
                    # print newjiage
                    z.append(newjiage)
                # print z
                else:
                    z.append(z[5])

    book = Workbook()
    for x in smulu:
        sheet1 = book.add_sheet(x)
        for y in needi:
            if x in y:
                sss = copy.deepcopy(y.get(x))
                for i in range(len(sss)):
                    for j in range(len(sss[i])):
                        # print 'xxxxx'
                        if is_chinese(sss[i][j]):

                            sss[i][j].encode('utf-8')

                        # elif not sss[i] and sss[i]!=0:
                        # 	print "空值",
                        elif is_num(sss[i][j]) == 1:
                            if math.modf(sss[i][j])[0] == 0 or sss[i][j] == 0:  # 获取数字的整数和小数
                                sss[i][j] = int(sss[i][j])  # 将浮点数化成整数

                        # print sss[i][j]
                        if j == 0:
                            try:
                                cccc = pandas.to_datetime(sss[i][j] - 25569, unit='d')
                                sss[i][j] = str(pandas.Period(cccc, freq='D'))
                            except:
                                pass
                        sheet1.write(i, j, sss[i][j])
    book.save('suoyou.xls')  # 存储excel
    book = xlrd.open_workbook('suoyou.xls')

# addmoney()

# print yi,er,san,si
# print yilist,'yi',len(yilist)
# print erlist,'er',len(erlist)
# print sanlist,'san',len(sanlist)
# print silist,'si',len(silist)
# print errortype,'errortype'
