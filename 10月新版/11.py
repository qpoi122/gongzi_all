from __future__ import division
import xlrd
import os
import math
from xlwt import Workbook, Formula
import xlrd
import copy
import pandas
import types
#
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


def add_name(name_list):
    for x in name_list:
        aa = dict.fromkeys(x, [])
        yield aa


# 不带颜色的读取
def load_file(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    file = (file_excel + '.xls')  # 文件名及中文合理性
    if not os.path.exists(file):  # 判断文件是否存在
        file = (file_excel + '.xlsx')
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    # print('suicce')


# [[],[]]
def load_file_with_twolist(file_name):
    # money为所有工资计算对应的资料
    file_list =[]
    load_file(file_name)

    Sheetname = workbook.sheet_names()
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
                    if math.modf(a[i])[0] == 0 or a[i] == 0:
                        a[i] = int(a[i])
                try:
                    a[i] = int(a[i])
                except:
                    pass
                mid.append(a[i])
            file_list.append(mid)
    return file_list


def load_zhangdan(file_name):
    #renming为每个分组的人员组成，zhangdan为整个账簿的信息
    renming = [[1], [2], [3], [4]]
    zhangdan = []


    load_file(file_name)

    Sheetname = workbook.sheet_names()
    for name in range(len(Sheetname)):
        table = workbook.sheets()[name]
        nrows = table.nrows
        title = table.row_values(1)
        for x in range(len(title)):
            if '.' in title[x]:
                gongxuleixing = title[x].split(u'.')[0]
                for y in renming:
                    if int(gongxuleixing) == y[0]:
                        zuhemingcheng = title[x].split(u'.')[1]
                        if u'/' in zuhemingcheng:
                            splitxingming = zuhemingcheng.split(u'/')
                            for z in splitxingming:
                                if z not in y:
                                    y.append(z)
                        else:
                            y.append(zuhemingcheng)

        for x in renming:
            del x[0]

        for n in range(2, nrows):
            a = table.row_values(n)
            mid = []
            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')

                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:
                        a[i] = int(a[i])
                try:
                    a[i] = a[i].strip()
                except:
                    pass

                mid.append(title[i])
                mid.append(a[i])
            zhangdan.append(mid)
    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')
    for i in range(len(zhangdan)):
        for j in range(len(zhangdan[i])):
            if is_chinese(zhangdan[i][j]):
                zhangdan[i][j].encode('utf-8')

            elif is_num(zhangdan[i][j]) == 1:
                if math.modf(zhangdan[i][j])[0] == 0 or zhangdan[i][j] == 0:
                    zhangdan[i][j] = int(zhangdan[i][j])
            sheet1.write(i, j, zhangdan[i][j])
    book.save('allmesg.xls')  # 存储excel
    book = xlrd.open_workbook('allmesg.xls')

    return renming, zhangdan
    print('')

# [[1,[]],[2,[]]]
def load_file_with_threelist(file_name):

    file_list = []
    load_file(file_name)

    Sheetname = workbook.sheet_names()
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):

            a = table.row_values(n)
            mid = []

            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')
                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:
                        a[i] = int(a[i])
                try:
                    a[i] = int(a[i])
                except:
                    pass
                mid.append(a[i])
            midd.append(mid)
        file_list.append(midd)
    return file_list


def mix_money_zhangdan(money, zhangdan, dictmesg_with_name):
    for x in zhangdan:
        for z in range(len(x)):
            # 提前获取各个信息
            pd = [u'客户编号', u'总成型号', u'数量', u'日期', u'塑料袋', u'小标贴', u'小内盒']
            if x[z] == u'总成型号':
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
                for h in range(len(dictmesg_with_name)):
                    if pipei_num[m] in dictmesg_with_name[h] and x[z + 1] != u'' and x[z + 1] != u' ' and type(
                            x[z + 1]) != type(u''):
                        # zong
                        zhongji = []
                        # xiao
                        mid = []
                        zhongji = copy.deepcopy(dictmesg_with_name[h].get(pipei_num[m]))

                        # 加上匹配到的项的日期
                        flgggggg = 0
                        a = z - 1
                        for p in range(0, 2):
                            if type(x[a]) == type(1):
                                mid.append(x[a])
                                flgggggg = 1
                                break
                            else:
                                a = a - 1
                        if flgggggg == 0:
                            mid.append(u'no tiqi')
                        mid.append(zcbh)
                        mid.append(khbh)

                        if type(x[z + 1]) != type(1):
                            print(z, x[z + 1], type(x[z + 1]), type(len(pipei_num)))
                        mid.append(len(pipei_num))
                        # 划分单价
                        shul = x[z + 1]
                        mid.append(shul)

                        zongmoney = 0
                        finded = 0
                        for j in range(len(money)):
                            # 将除了第四道的工序都匹配金钱,没有匹配到的钱为0
                            if int(money[j][0]) == h + 1 and money[j][1] == x[3] and int(money[j][0]) != 4:
                                try:
                                    zongmoney = money[j][2] * shul / len(pipei_num)
                                except:
                                    pass
                                break
                        mid.append(zongmoney)
                        # 如果是第四种加上√ 信息
                        if h == 3:
                            mid.append(sld)
                            mid.append(xbt)
                            mid.append(xnh)
                        zhongji.append(mid)
                        dictmesg_with_name[h][pipei_num[m]] = zhongji
    return dictmesg_with_name


def check_zhangdan_with_taizhang(dictmesg_with_name, taizhang):
    smulu = []
    for x in dictmesg_with_name:
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
        for y in dictmesg_with_name:
            if x in y:
                zmsj = copy.deepcopy(y.get(x))
                for z in zmsj:
                    del z[3]
                    z.pop()
                    if x in name4:
                        del z[4:]
                break

        tj = []
        for z in range(len(taizhang)):
            if x in taizhang[z][0]:
                tj = copy.deepcopy(taizhang[z])
                break

        if tj != []:
            for m in range(len(zmsj)):
                flag = 0
                for n in range(len(tj)):
                    if zmsj[m][1] == tj[n][1] and zmsj[m][2] == tj[n][2] and zmsj[m][3] == tj[n][3] and tj[n][-1] != u'o':
                        try:
                            if 0 <= abs(zmsj[m][0] - tj[n][0]) <= 5 or 365 <= abs(zmsj[m][0] - tj[n][0]) <= 371:
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

        titles = x
        sheet1 = book1.add_sheet(titles)
        for i in range(1, len(tj)):
            for j in range(len(tj[i])):
                if is_chinese(tj[i][j]):

                    tj[i][j].encode('utf-8')
                elif is_num(tj[i][j]) == 1:
                    if math.modf(tj[i][j])[0] == 0 or tj[i][j] == 0:
                        tj[i][j] = int(tj[i][j])

                if j == 0:
                    try:
                        cccc = pandas.to_datetime(tj[i][j] - 25569, unit='d')
                        tj[i][j] = str(pandas.Period(cccc, freq='D'))
                    except:
                        pass
                sheet1.write(i - 1, j, tj[i][j])
    book1.save('no danzi.xls')  # 存储excel
    book1 = xlrd.open_workbook('no danzi.xls')

    book2 = Workbook()
    for x in smulu:
        sheet1 = book2.add_sheet(x)
        for y in lackzm:
            if x in y:
                sss = copy.deepcopy(y.get(x))
                for i in range(len(sss)):
                    for j in range(len(sss[i])):
                        if is_chinese(sss[i][j]):

                            sss[i][j].encode('utf-8')

                        elif is_num(sss[i][j]) == 1:
                            if math.modf(sss[i][j])[0] == 0 or sss[i][j] == 0:
                                sss[i][j] = int(sss[i][j])
                        if j == 0:
                            try:
                                cccc = pandas.to_datetime(sss[i][j] - 25569, unit='d')
                                sss[i][j] = str(pandas.Period(cccc, freq='D'))
                            except:
                                pass
                        sheet1.write(i, j, sss[i][j])
    book2.save('no taizhang.xls')  # 存储excel
    book2 = xlrd.open_workbook('no taizhang.xls')

    return smulu


def fuzzy_matching(nodanzi, notaizhang):
    pijiaoout = []
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

    book5 = Workbook()
    for z in range(len(pijiaoout)):
        sheet1 = book5.add_sheet(pijiaoout[z][0])
        for i in range(1, len(pijiaoout[z])):
            for j in range(len(pijiaoout[z][i])):
                if is_chinese(pijiaoout[z][i][j]):
                    pijiaoout[z][i][j].encode('utf-8')
                elif is_num(pijiaoout[z][i][j]) == 1:
                    if math.modf(pijiaoout[z][i][j])[0] == 0 or pijiaoout[z][i][j] == 0:
                        pijiaoout[z][i][j] = int(pijiaoout[z][i][j])
                sheet1.write(i, j, pijiaoout[z][i][j])
    book5.save('mohu.xls')
    book5 = xlrd.open_workbook('mohu.xls')

    # #模糊匹配失败的，查看前后是否有同型号，同客户，同数量的也算漏记
    nodanzicopy = copy.deepcopy(nodanzi)

    for x in range(len(nodanzi)):
        for y in range(len(nodanzicopy)):
            for z in range(1, len(nodanzi[x])):
                for i in range(1, len(nodanzicopy[y])):
                    if len(nodanzi[x]) > 1 and len(nodanzicopy[y]) > 1:
                        if nodanzi[x][z][1:] == nodanzicopy[y][i][1:] and nodanzi[x][0] != nodanzicopy[y][0]:
                            if (not (nodanzi[x][0] in renming[0] and nodanzicopy[y][0] in renming[0])):
                                if nodanzi[x][z][-1] != 'ok' and nodanzicopy[x][z][-1] != 'ok':
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

                    elif is_num(nodanzi[z][i][j]) == 1:
                        if math.modf(nodanzi[z][i][j])[0] == 0 or nodanzi[z][i][j] == 0:
                            nodanzi[z][i][j] = int(nodanzi[z][i][j])
                    sheet1.write(line, j, nodanzi[z][i][j])

        line = -1
        for i in range(1, len(notaizhang[z])):
            if notaizhang[z][i][-1] != u'ok':
                line = line + 1
                for j in range(len(notaizhang[z][i])):
                    if is_chinese(notaizhang[z][i][j]):
                        notaizhang[z][i][j].encode('utf-8')
                    elif is_num(notaizhang[z][i][j]) == 1:
                        if math.modf(notaizhang[z][i][j])[0] == 0 or notaizhang[z][i][j] == 0:
                            notaizhang[z][i][j] = int(notaizhang[z][i][j])
                    sheet1.write(line, j + 6, notaizhang[z][i][j])
    book6.save('lackmohu.xls')
    book6 = xlrd.open_workbook('lackmohu.xls')


def specaldel_with_lastprocess(dictmesg_with_name, smulu, money):
    # 对第四道进行单独处理,先匹配名字一样的,没有再去按老方法处理
    for x in smulu:
        if x in dictmesg_with_name[3]:
            for y in dictmesg_with_name[3].get(x):
                for z in range(len(money)):
                    if money[z][0] == u'4' and y[1] == money[z][1]:
                        duigeshu = len(y[7])
                        jiage = ((duigeshu - 1) * 0.005 + money[z][2]) / y[3]
                        y[5] = jiage * y[4]

    for x in smulu:
        if x in dictmesg_with_name[3]:
            for y in dictmesg_with_name[3].get(x):
                flag = 0
                for z in range(len(money)):
                    if money[z][0] == u'5' and y[5] == 0:
                        # 内盒一样并且带-的
                        if y[8] == money[z][3] and u'-' == money[z][1] and type(y[1]) == type(u'a'):
                            jiage = money[z][4] / y[3]
                            y[5] = jiage * y[4]
                            flag = 1
                            break
                        # 所有带几盒几个的,但是不带-的
                        elif u'个' in y[7] or u'只' in y[7]:
                            if (type(y[1]) == type(u'a') and u'-' not in y[1]) or type(y[1]) == type(1):
                                if y[6] == money[z][1] and y[8] == money[z][3]:
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
                                        break
                        # 带排/数的
                        elif (type(y[1]) == type(u'a') and u'-' not in y[1] and (y[8] == u'排' or y[8] == u'数')) or (
                                type(y[1]) == type(1) and (y[8] == u'排' or y[8] == u'数')):
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
                                break
    return dictmesg_with_name


def fill_price_difference(dictmesg_with_name, money):
    # 补差价
    tijia = []
    for x in range(len(money)):
        if money[x][0] == u'6':
            tijia.append(money[x])

    for x in range(len(tijia)):
        for y in renming[tijia[x][1] - 1]:
            for z in dictmesg_with_name[tijia[x][1] - 1].get(y):
                if u'' in tijia[x]:
                    flag = 0
                    for j in range(len(tijia[x])):
                        if tijia[x][j] == tijia[x][1]:
                            if z[4] <= tijia[x][j + 1]:
                                newjiage = z[5] * (1 + tijia[x][j + 2])
                                flag = 1
                                z.append(newjiage)
                                break
                    if flag == 0:
                        z.append(z[5])

                elif z[4] <= tijia[x][2]:
                    newjiage = z[5] * (1 + tijia[x][3])
                    z.append(newjiage)
                else:
                    z.append(z[5])

    book = Workbook()
    for x in smulu:
        sheet1 = book.add_sheet(x)
        for y in dictmesg_with_name:
            if x in y:
                sss = copy.deepcopy(y.get(x))
                for i in range(len(sss)):
                    for j in range(len(sss[i])):
                        if is_chinese(sss[i][j]):
                            sss[i][j].encode('utf-8')
                        elif is_num(sss[i][j]) == 1:
                            if math.modf(sss[i][j])[0] == 0 or sss[i][j] == 0:
                                sss[i][j] = int(sss[i][j])

                        # print sss[i][j]
                        if j == 0:
                            try:
                                cccc = pandas.to_datetime(sss[i][j] - 25569, unit='d')
                                sss[i][j] = str(pandas.Period(cccc, freq='D'))
                            except:
                                pass
                        sheet1.write(i, j, sss[i][j])
    book.save('suoyou.xls')
    book = xlrd.open_workbook('suoyou.xls')

def load_lack_zhangdan(file_name):
    file = []
    load_file(file_name)

    Sheetname = workbook.sheet_names()
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        midd = []
        midd.append(ttype)
        for n in range(nrows):
            a = table.row_values(n)
            mid = []
            if a[-1] != u'o':
                for i in range(len(a)):

                    if is_chinese(a[i]):
                        a[i].encode('utf-8')
                    elif is_num(a[i]) == 1:
                        if math.modf(a[i])[0] == 0 or a[i] == 0:
                            a[i] = int(a[i])
                    mid.append(a[i])
                midd.append(mid)
        file.append(midd)
    return file


def change_taizhang_order(taizhang):
    for x in taizhang:
        for y in x:
            if type(y) != type('s'):
                a = y[2]
                y[2] = y[1]
                y[1] = a
    return taizhang


def clear_file():
    file_list= ['no danzi.xls', 'no taizhang.xls', 'allmesg.xls']
    for file in file_list:
        if os.path.exists(file):
            os.remove(file)
        else:
            pass

if __name__ == "__main__":
    money = load_file_with_twolist('newmy')
    renming, zhangdan = load_zhangdan('12345')
    taizhang_1 = load_file_with_threelist('taizhang')
    taizhang = change_taizhang_order(taizhang_1)
    #创建一个带每个人信息抬头的空字典
    dictmesg_with_name = []
    add_name = add_name(renming)
    for o in add_name:
        dictmesg_with_name.append(o)

    dictmesg_with_name = mix_money_zhangdan(money, zhangdan, dictmesg_with_name)
    smulu = check_zhangdan_with_taizhang(dictmesg_with_name, taizhang)
    nodanzi = load_lack_zhangdan('no danzi')
    notaizhang = load_lack_zhangdan('no taizhang')
    fuzzy_matching(nodanzi, notaizhang)
    dictmesg_with_name = specaldel_with_lastprocess(dictmesg_with_name, smulu, money)
    fill_price_difference(dictmesg_with_name, money)
    clear_file()