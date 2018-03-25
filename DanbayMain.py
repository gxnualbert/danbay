#!usr/bin/env python
#-*- coding:utf-8 _*-
"""
@author:albert.chen
@file: DanbayMain.py
@time: 2018/03/25/20:44
"""

#!usr/bin/env python
#-*- coding:utf-8 -*-
"""
@author:albert.chen
@file: login.py
@time: 2017/12/27/14:47
"""

import requests,json
import MySQLdb,time,os
import xlrd,xlwt,datetime
from xlutils.copy import copy
from xlwt import *
import ConfigParser
import sys
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)

def renameFile(fileName):
    '''
    传进来的文件要包含路径和名字
    :param fileName:
    :return:
    '''
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls",unicode(fileName, "utf-8") + u'_门锁密码容量_'+unicode(a,"utf-8")+u'.xls')
    # f.save(unicode(fileName, "utf-8") + u'_门锁密码容量_'+unicode(a,"utf-8")+u'.xls')
    # os.rename("test.xls","lock_Info_"+a+".xls")

'''
设置单元格格式
'''
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style
'''
操作Excel 表格
'''
def redStyle():
    red_style = xlwt.XFStyle()  # 初始化样式
    pattern = Pattern()  # 创建一个模式
    pattern.pattern = Pattern.SOLID_PATTERN  # 设置其模式为实型
    pattern.pattern_fore_colour = 2
    red_style.pattern = pattern
    return  red_style
def dateStyle():
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyy/mm/dd'
    return date_format
def generateExcel():
    f = xlwt.Workbook(encoding='utf-8') #创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'门锁密码统计',cell_overwrite_ok=True) #创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'门锁ID',u'密码已使用数']
    #生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0,len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 18
        sheet1.col(2).width = 256 * 15
        sheet1.col(3).width = 256 * 15
        sheet1.col(4).width = 256 * 15
        sheet1.col(5).width = 256 * 15
        sheet1.col(6).width = 256 * 15
        sheet1.col(7).width = 2+56 * 50
        sheet1.col(8).width = 256 * 15
        sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True))
    f.save("test.xls")
    # f.save(unicode(fileName, "utf-8") + u'_门锁密码容量.xls')  # 保存文件
    # f.save(unicode(fileName, "utf-8") + u'_门锁密码容量.xls')
def meterExcel():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    sheet1 = f.add_sheet(u'电表电量记录', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间号', u'冻结时间', u'电表读数']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 30
        sheet1.col(1).width = 256 * 18
        sheet1.col(2).width = 256 * 15
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")
def waterExcel():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    sheet1 = f.add_sheet(u'水表水量记录', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间号', u'用量']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 30
        sheet1.col(1).width = 256 * 18
        sheet1.col(2).width = 256 * 15
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("water.xls")
def generateLockInfoExcel():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'门锁在线离线状态', cell_overwrite_ok=True)  # 创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'房间地址', u'门锁状态', u'门锁设备ID',u'门锁Mac',u'中控状态',u'中控设备ID',u'中控Mac']
    # 生成第一行,并设置单元格长度，只需设置一次即可状态
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 60
        sheet1.col(1).width = 256 * 10
        sheet1.col(2).width = 256 * 50
        sheet1.col(3).width = 256 * 20
        sheet1.col(4).width = 256 * 10
        sheet1.col(5).width = 256 * 50
        sheet1.col(6).width = 256 * 40
        # sheet1.col(7).width = 2 + 56 * 50
        # sheet1.col(8).width = 256 * 15
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")
def generateLockPwdInfoExcel():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'碧桂园门锁预置密码表', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间地址', u'设备ID',u'预置租客密码', u'预置临时密码1',u'预置临时密码2',u'预置临时密码3']
    # 生成第一行,并设置单元格长度，只需设置一次即可状态
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 60
        sheet1.col(1).width = 256 * 20
        sheet1.col(1).width = 256 * 15
        sheet1.col(2).width = 256 * 16
        sheet1.col(3).width = 256 * 16
        sheet1.col(4).width = 256 * 16
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")

def generateMeterExcel():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'水电表信息统计', cell_overwrite_ok=True)  # 创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'房间地址', u'设备类型', u'设备状态', u'houseID',u'设备ID',u'关联中控地址',u'中控mac',u'中控状态',u'中控device id']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 18
        sheet1.col(2).width = 256 * 15
        sheet1.col(3).width = 256 * 15
        sheet1.col(4).width = 256 * 50
        sheet1.col(5).width = 256 * 80
        sheet1.col(6).width = 256 * 80
        sheet1.col(7).width = 256 * 80


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")

def generateCaiJiQi():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'采集器水电表读数信息', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间地址', u'设备类型',u'水电表在线状态',u'采集器ID', u'水电表标号', u'上报记录数',u'中控上报payLoadString'u'上报时间']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 12
        sheet1.col(1).width = 256 * 19
        sheet1.col(1).width = 256 * 18
        sheet1.col(2).width = 256 * 15
        sheet1.col(3).width = 256 * 30
        sheet1.col(4).width = 256 * 20
        # sheet1.col(5).width = 256 * 80
        # sheet1.col(6).width = 256 * 80
        # sheet1.col(7).width = 256 * 80


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")

def generateShuiDianReading():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'水电表表头读数', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间地址', u'设备类型',u'在线状态',u'设备表号',u'设备当前读数']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 12
        sheet1.col(2).width = 256 * 15
        sheet1.col(3).width = 256 * 20
        sheet1.col(4).width = 256 * 15


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")

def generateRiZhiBiao():
    f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    '''
    创建第一个sheet:
        sheet1
    '''
    sheet1 = f.add_sheet(u'日志表信息', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'房间地址', u'设备类型',u'采集器DevID',u'设备表号',u'PayLoad', u'上报时间']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 12
        sheet1.col(2).width = 256 * 30
        sheet1.col(3).width = 256 * 19
        sheet1.col(4).width = 256 * 50
        sheet1.col(5).width = 256 * 25
        # sheet1.col(3).width = 256 * 30
        # sheet1.col(4).width = 256 * 20
        # sheet1.col(5).width = 256 * 80
        # sheet1.col(6).width = 256 * 80
        # sheet1.col(7).width = 256 * 80


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("rizhi.xls")
def generateCenterControlExcel():
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'房间地址', u'中控id', u'中控状态', u'该中控下的在线设备数', u'该中控设备总数',u'中控deviceID',u'中控Mac地址',u'中控版本']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 50
        sheet1.col(1).width = 256 * 15
        sheet1.col(2).width = 256 * 15
        sheet1.col(3).width = 256 * 25
        sheet1.col(4).width = 256 * 15
        sheet1.col(5).width = 256 * 25
        sheet1.col(6).width = 256 * 25
        sheet1.col(7).width = 256 * 15
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")

def generateCenterControlinfo():
    # centerContronDevmacAddress = devicersp["macAddress"]
    # centerContronDevVersion = devicersp["deviceModel"]
    # centerContronDevID = devicersp["deviceId"]
    # centerContronDevaddress = devicersp["address"]
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'中控MAC地址', u'中控版本', u'中控设备ID', u'中控地址',u'中控在线离线状态']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 20
        sheet1.col(1).width = 256 * 23
        sheet1.col(2).width = 256 * 40
        sheet1.col(3).width = 256 * 150


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")
def read_excel():

    #文件位置

    ExcelFile=xlrd.open_workbook(u'安徽省合肥市蜀山区稻香路与山湖路交口向西100米_水电表状态_2018-02-27_23_29_07.xls')

    #获取目标EXCEL文件sheet名

    print ExcelFile.sheet_names()[0]

    #------------------------------------

    #若有多个sheet，则需要指定读取目标sheet例如读取sheet2

    #sheet2_name=ExcelFile.sheet_names()[1]

    #------------------------------------

    #获取sheet内容【1.根据sheet索引2.根据sheet名称】

    sheet=ExcelFile.sheet_by_index(0)

    # sheet=ExcelFile.sheet_by_name('TestCase002')

    #打印sheet的名称，行数，列数

    print sheet.name,sheet.nrows,sheet.ncols
    #
    # #获取整行或者整列的值
    #
    # rows=sheet.row_values(2)#第三行内容
    #
    # cols=sheet.col_values(1)#第二列内容
    #
    # print cols,rows
    #
    # #获取单元格内容
    #
    print sheet.cell(1,0).value.encode('utf-8')
    #
    # print sheet.cell_value(1,0).encode('utf-8')
    #
    # print sheet.row(1)[0].value.encode('utf-8')

    #打印单元格内容格式


def generateCenterControlInfoAtDeviceCenter():
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
    # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
    row0 = [u'中控MAC地址', u'中控版本',  u'中控在线离线状态',u'中控地址']
    # 生成第一行,并设置单元格长度，只需设置一次即可
    for i in range(0, len(row0)):
        sheet1.col(0).width = 256 * 20
        sheet1.col(1).width = 256 * 23
        sheet1.col(2).width = 256 * 40
        sheet1.col(3).width = 256 * 150


        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    f.save("test.xls")
def writeExcel(rowIndex,colIndex,cellValue,userStyle=1):
    '''
    该函数主要功能是实现往已存在的Excel表格中写入数据
    :param workBookNmae: Excel 表名
    :param rowIndex: 行索引
    :param colIndex: 列索引
    :param cellValue: 单元格的值
    :return:
    '''

    rb = xlrd.open_workbook("test.xls",formatting_info=True)
    # sheet=data.sheet_by_name(u'门锁密码统计')
    wb = copy(rb)
    # write(1, 0, "test")，第一个是行索引，第二个是列索引，第三个是单元格的值
    if userStyle==1:
        wb.get_sheet(0).write(rowIndex,colIndex,cellValue)
    else:
        userStyle=redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue,userStyle)
    wb.save(u'test.xls')

def writeRiZhi(rowIndex,colIndex,cellValue,userStyle=1):
    '''
    该函数主要功能是实现往已存在的Excel表格中写入数据
    :param workBookNmae: Excel 表名
    :param rowIndex: 行索引
    :param colIndex: 列索引
    :param cellValue: 单元格的值
    :return:
    '''

    rb = xlrd.open_workbook("rizhi.xls",formatting_info=True)
    # sheet=data.sheet_by_name(u'门锁密码统计')

    # write(1, 0, "test")，第一个是行索引，第二个是列索引，第三个是单元格的值
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
    # rb = xlrd.open_workbook("water.xls", formatting_info=True)
    wb = copy(rb)
    if userStyle == 1:
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    else:
        userStyle = redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
    wb.save(u'rizhi.xls')

def waterWriteExcel(rowIndex,colIndex,cellValue,userStyle=1):
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy/mm/dd'
    rb = xlrd.open_workbook("water.xls", formatting_info=True)
    wb = copy(rb)
    if userStyle == 1:
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    else:
        userStyle = redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
    wb.save(u'water.xls')

def TwowaterWriteExcel(rowIndex,colIndex,cellValue,userStyle=1):
    # date_format = xlwt.XFStyle()
    # date_format.num_format_str = 'yyyy/mm/dd'
    rb = xlrd.open_workbook("water.xls", formatting_info=True)
    wb = copy(rb)
    if userStyle == 1:
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    else:
        # userStyle = redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    wb.save(u'water.xls')

def meterWriteExcel(rowIndex,colIndex,cellValue,userStyle=1):
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy/mm/dd'
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    if userStyle == 1:
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    else:
        userStyle = redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
    wb.save(u'test.xls')

def TwometerWriteExcel(rowIndex,colIndex,cellValue,userStyle=1):
    date_format = xlwt.XFStyle()
    # date_format.num_format_str = 'yyyy/mm/dd'
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    if userStyle == 1:
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
    else:
        userStyle = redStyle()
        wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
    wb.save(u'test.xls')

def dbOperation(sql,db='danbay_device'):
    conn = MySQLdb.connect(
        host='rm-wz916f30z77a773rdo.mysql.rds.aliyuncs.com',
        port=3306,
        user='danbay_read',
        passwd='LoveDanbayNow@',
        db=db,
        charset="utf8"
    )
    cur = conn.cursor()
    cur.execute(sql)
    results = cur.fetchall()

    cur.close()
    conn.commit()
    conn.close()
    return results

def getPwdCountsInPre(deviceID):
    '''
    获取预置密码表的密码数量，只有租客喝临时密码才有预置密码
    :param sql: 需要执行的sql语句
    :return: 返回一个密码总数的字典
    '''
    # sql = "SELECT psw_alias,psw_type from lock_pre_password WHERE dev_id='c2c60e1a0e70d71b9010718f213a8f13'"
    sql = "SELECT psw_alias,psw_type from lock_pre_password WHERE dev_id="+ "\'" + deviceID + "\'" + "and delete_state !=1;"
    result=dbOperation(sql)
    renter_pwd=0
    tmp_pwd=0
    for row in result:
        if row[1]=="3":
            renter_pwd=renter_pwd+1
        if row[1]=="0":
            tmp_pwd=tmp_pwd+1
    # print renter_pwd,tmp_pwd
    pre_pwd_count={}
    pre_pwd_count["renter_pwd"]=renter_pwd
    pre_pwd_count["tmp_pwd"]=tmp_pwd
    return pre_pwd_count
def getPwdCountsInNormal(deviceID):
    '''
    获取正常密码表的密码数量
    :param sql: 需要执行的sql语句
    :return: 返回密码总数的字典
    '''
    device_info_sql = "SELECT id FROM device_info WHERE deviceId=" + "\'" + deviceID + "\'" + ";"
    result = dbOperation(device_info_sql)
    id=0
    for row in result:
       id=row[0]
    device_pwd_info_sql = "SELECT pwdType,pwdAlias from device_pwd_info WHERE deviceInfo=" + "\'" + str(id) + "\'" + ";"
    pwdresult=dbOperation(device_pwd_info_sql)
    housekeeper=0
    admin=0
    renter=0
    tmp=0
    for pwdinfo in pwdresult:
        if pwdinfo[0]=="0":
            tmp = tmp + 1
        elif pwdinfo[0]=="1":
            admin=admin+1
        elif pwdinfo[0] == "2":
            housekeeper=housekeeper+1
        elif pwdinfo[0] == "3":
            renter=renter+1
    # print tmp,admin,housekeeper,renter
    nor_pwd_count={}
    nor_pwd_count["tmp"]=tmp
    nor_pwd_count["admin"] =admin
    nor_pwd_count["housekeeper"] =housekeeper
    nor_pwd_count["renter"] =renter
    # print nor_pwd_count
    return nor_pwd_count
def getHomeAddress(ID):
    # 'SELECT homeAddress FROM house_info WHERE ID='24052''
    device_info_sql = "SELECT homeAddress FROM house_info WHERE ID=" + "\'" + str(ID) + "\'" + ";"
    result = dbOperation(device_info_sql)
    homeaddress=result[0]
    return homeaddress[0]
    # print  homeaddress[0]
def getHouseID(devuiceID):
    # 'SELECT * from energy_device where deviceId='a1bb95ed70878d2d729b204e140eff4e''
    houseIDSql="SELECT houseId from energy_device where deviceId="+ "\'" + str(devuiceID) + "\'" + ";"
    result = dbOperation(houseIDSql)
    houseID = result[0]
    # print houseID[0]
    return houseID[0]
def getCookies(username,password):
    loginurl = 'http://www.danbay.cn/system/goLoginning'
    # payload = {'mc_username': 'admin', 'mc_password': 'Danbay@20171214!', 'rememberMe': ""}
    payload = {'mc_username': username, 'mc_password':password, 'rememberMe': ""}
    r = requests.post(loginurl, data=payload)
    return r.cookies

def getDeviceListAllWithCorrectInfo(username,password,address):
    generateExcel()
    payload = {'pageNo': '1', 'pageSize': '10', 'likeStr':address,"detailAddress":"","floor":"","spaceType":"","userId":"1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck=getCookies(username,password)
    # print ck
    r = requests.post(host, data=payload,cookies=ck)
    rsp=json.loads(r.text)
    rsp=rsp["result"]
    rowcount=rsp["pageCount"]
    print "返回的页数为：",rowcount
    # 为了调试方便，暂时设置为10
    # rowcount=10
    payload = {'pageNo': '1', 'pageSize':str(rowcount) , 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    rsp=rsp["resultList"]
    lockInfoDic={}
    roomInfoList = []
    for k in range(len(rsp)):
        i=rsp[k]
        houseid=i["hosueInfo"]["id"]
        req="http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s"%houseid
        r = requests.get(req,cookies=ck)
        lock_rsp=r.text
        lock_rsp=json.loads(lock_rsp,encoding='utf-8')
        lock_rsp=lock_rsp["result"]
        if lock_rsp["lockList"]:
            # 先根据devid 获取门锁的预置密码，然后再根据devid找到id，再用id找到正常密码
            # print lock_rsp["roomName"],lock_rsp["lockList"][0]["deviceId"]
            lockInfoDic[lock_rsp["lockList"][0]["deviceId"]]=lock_rsp["roomName"]
            deviceID=lock_rsp["lockList"][0]["deviceId"]
            # if "60280c1cc487c835de88f150ec65c173"==lock_rsp["lockList"][0]["deviceId"]:
            #     print "now stat to see the value"
            #     print "ha hahah ha"
            nor_pwd_count=getPwdCountsInNormal(deviceID)
            pre_pwd_count=getPwdCountsInPre(deviceID)
            # pre_pwd_count["renter_pwd"] = renter_pwd
            # pre_pwd_count["tmp_pwd"] = tmp_pwd
            # print i
            # u'管理员密码个数', u'管家密码个数', u'租客密码个数', u'临时密码个数', u'预置租客个数', u'预置临时个数', u'密码已使用数'
            # 计算密码已使用个数
            admincount=nor_pwd_count["admin"]
            housekeepercount=nor_pwd_count["housekeeper"]
            renterCount=nor_pwd_count["renter"]
            tmpCount=nor_pwd_count["tmp"]
            renterPreCount=pre_pwd_count["renter_pwd"]
            tmpPreCount=pre_pwd_count["tmp_pwd"]
            locktotalCount=int(admincount)+int(housekeepercount)+int(renterCount)+int(tmpCount)+int(renterPreCount)+int(tmpPreCount)
            rowTupple=(lock_rsp["roomName"],nor_pwd_count["admin"],nor_pwd_count["housekeeper"],nor_pwd_count["renter"],nor_pwd_count["tmp"],pre_pwd_count["renter_pwd"],pre_pwd_count["tmp_pwd"],locktotalCount,lock_rsp["lockList"][0]["deviceId"])
            roomInfoList.append(rowTupple)
    aa=sorted(roomInfoList,key=lambda room:room[7],reverse=True)
    # print aa
    # aa是排序后的结果，排序之后，直接写进Excel表格中，上面的sorted 函数，最后加个reverse=True表示
    # 降序排列
    for roomRowIndex in range(len(aa)):
        roomRow=aa[roomRowIndex]
        writeExcel(roomRowIndex+1,0,roomRow[0])
        writeExcel(roomRowIndex+1,1,roomRow[1])
        writeExcel(roomRowIndex+1,2,roomRow[2])
        writeExcel(roomRowIndex+1,3,roomRow[3])
        writeExcel(roomRowIndex+1,4,roomRow[4])
        writeExcel(roomRowIndex+1,5,roomRow[5])
        writeExcel(roomRowIndex+1,6,roomRow[6])
        writeExcel(roomRowIndex+1,7,roomRow[8])
        writeExcel(roomRowIndex+1,8,roomRow[7])

    print address[0]
    renameFile(address[0])

    end=datetime.datetime.now()



def getPayload(pageNo,address):
    payload = {'pageNo': pageNo, 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    return payload

def getCenterControlPayload(pageNo,address):
    payload = {'pageNo': pageNo, 'pageSize': '8', "status": "",'likeStr': address, "isNewVersion": ""}
    return payload

def getJson(ck,houseid):
    '''
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
            <title>500服务器出错</title>
        </head>
        <body>
            <h4 align="center">服务器出错了，请稍后访问或者联系管理员！</h4>
        </body>
        </html>
        在20180109 这段时间，向服务器请求数据的时候，会随机报这种错误，导致程序中断，已经在群里反馈
        为了后面不中断程序运行，我这里直接try  catch 捕获他
    :param ck:
    :param houseid:
    :return:
    '''
    for i in range(10):
        req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
        r = requests.get(req, cookies=ck)
        devicersp = r.text
        # a = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        # print a, houseid
        # print r,devicersp
        # return devicersp

        # print devicersp
        # //由于devicersp 是unicode 类型，所以要转一下if u"服务器出错了" in devicersp:
        if u"服务器出错了" in devicersp:
            print "通过HouseId 获取设备信息失败，尝试次数为：%s" % i
            # 如果服务器报错了，休眠1分钟，然后再请求
            time.sleep(30)
        else:
            return devicersp

def getCenterControlJson(ck,centerControlID):
    for i in range(10):
        centerControlURL = "http://www.danbay.cn/system/centerControl/getCenterControlDetail?id=%s"%centerControlID
        r = requests.post(centerControlURL, cookies=ck)
        # req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
        # r = requests.get(req, cookies=ck)
        devicersp = r.text
        # //由于devicersp 是unicode 类型，所以要转一下if u"服务器出错了" in devicersp:
        if u"服务器出错了" in devicersp:
            print "通过中控ID 获取设备信息失败，尝试次数为：%s" % i
            #如果服务器报错了，休眠1分钟，然后再请求
            time.sleep(30)
        else:
            return devicersp
def getPageCount(username,password,address):
    payload = getPayload("1", address)
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pageCount = rsp["pageCount"]
    print "总共的页数是： ",pageCount
    return pageCount

def centerControlGetPageCount(username,password,address):
    payload = getCenterControlPayload("1", address)
    host = "http://www.danbay.cn/system/centerControl/findCenterControlSetting"
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]["result"]
    pageCount = rsp["pageCount"]

    print "总共的页数是:", pageCount
    return pageCount
def getCenterControlpageCount():
    payload = getPayload("1", address)
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pageCount = rsp["pageCount"]
    print pageCount
    return pageCount

def getAmmeterDeviceId(username,password,address):
    # excelRows记录当前Excel写到了多少行
    excelRows=1
    generateMeterExcel()
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount=getPageCount(username,password,address)

    # pageCount=46
    for index in range(pageCount):
        pageNo=index+1
        payload=getPayload(pageNo,address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        print "页面%s,响应记录是%s"%(str(pageNo),str(len(rsp)))
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp=getJson(ck,houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            meterList=devicersp["meterList"]

            #将水表或电表的状态写入表格
            for deviceIndex in range(len(meterList)):
                meter=meterList[deviceIndex]
                homeaddress=i["hosueInfo"]["homeAddress"]
                meterType=meter["meterType"]
                subType=meter["subType"]
                meterID=meter["deviceId"]
                elecmeterID=meter["id"]
                meterStatus=meter["onlineStatus"]
                aa=getHouseID(meterID)
                writeExcel(excelRows, 0, homeaddress)
                # 设备类型，水保还是电表
                if meterType=="0":
                    if subType:
                        if subType=="1":
                            writeExcel(excelRows, 1, u"热水")
                        elif subType=="0":
                            writeExcel(excelRows, 1, u"冷水")
                    else:
                        writeExcel(excelRows, 1, u"水表")
                elif meterType=="1":
                    writeExcel(excelRows, 1, u"电表")
                # 在线状态，离线还是在线
                if meterStatus=="1":
                    writeExcel(excelRows, 2, u"离线",2)
                elif meterStatus=="0":
                    writeExcel(excelRows, 2, u"在线")
                # houseID，如果不为空就写进去
                if aa:
                    writeExcel(excelRows, 3, aa)
                writeExcel(excelRows,4,meterID)
                #根据deviceid找到对应的中控，并写入进去
                getmeterControlIdSql="SELECT centerControlId FROM energy_device WHERE deviceId="+ "\'" + meterID + "\'" + ";"
                centerCotrolID=dbOperation(getmeterControlIdSql)[0][0]
                getmeterControlAddress="SELECT address FROM center_control WHERE id="+ "\'" + str(centerCotrolID) + "\'" + ";"
                meterControlAddress=dbOperation(getmeterControlAddress)
                # 水电表关联的中控的地址信息

                writeExcel(excelRows, 5, meterControlAddress[0][0])
                #中控deviceid
                ControlAddress = "SELECT deviceId,macAddress,online FROM center_control WHERE id=" + "\'" + str(centerCotrolID) + "\'" + ";"
                centID=dbOperation(ControlAddress)
                # print centID[0][0],centID[0][1]

                writeExcel(excelRows, 6, centID[0][1])
                # 中控在线离线状态
                if centID[0][2]==0:
                     writeExcel(excelRows, 7, u'在线')
                elif centID[0][2]==1:
                     writeExcel(excelRows, 7, u'离线',2)

                writeExcel(excelRows, 8, centID[0][0])


                print "finish: "+homeaddress
                excelRows=excelRows+1

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表状态_' + unicode(a, "utf-8") + u'.xls')


def getCountCaijiqi(username,password,address):
    # excelRows记录当前Excel写到了多少行
    excelRows=1
    generateCaiJiQi()
    generateRiZhiBiao()
    rizhiRow=1
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount=getPageCount(username,password,address)

    # pageCount=46
    for index in range(pageCount):
        pageNo=index+1
        payload=getPayload(pageNo,address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        print "页面%s,响应记录是%s"%(str(pageNo),str(len(rsp)))
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp=getJson(ck,houseid)
            try :
                devicersp = json.loads(devicersp, encoding='utf-8')
            except:
                print "no json coulld be decode!!!!!!!!!!!!!!!"
                print devicersp
            devicersp = devicersp["result"]
            meterList=devicersp["meterList"]
            #将水表或电表的状态写入表格
            for deviceIndex in range(len(meterList)):
                meter=meterList[deviceIndex]
                homeaddress=i["hosueInfo"]["homeAddress"]
                meterType=meter["meterType"]
                subType=meter["subType"]
                meterID=meter["deviceId"]
                elecmeterID=meter["id"]
                meterStatus=meter["onlineStatus"]
                aa=getHouseID(meterID)
                 # "SELECT * FROM energy_day_consumption WHERE energyDevice=11510"
                # recoredCount="SELECT * FROM energy_day_consumption WHERE energyDevice=" + "\'" + str(elecmeterID) + "\'" + " ORDER BY readTime DESC LIMIT 11;"
                recoredCount="SELECT * FROM energy_day_consumption WHERE energyDevice=" + "\'" + str(elecmeterID) + "\'" + " ORDER BY readTime DESC;"
                # "SELECT * FROM energy_day_consumption WHERE energyDevice=11510 ORDER BY readTime DESC LIMIT 10;"
                # if len(dbOperation(recoredCount)) < 10:
                    # print "energy_day_consumption的记录数量为：",len(dbOperation(recoredCount))
                writeExcel(excelRows, 0, homeaddress)
                if meterType == "0":
                    if subType:
                        if subType == "1":
                            writeExcel(excelRows, 1, u"热水")
                        elif subType == "0":
                            writeExcel(excelRows, 1, u"冷水")
                elif meterType == "1":
                    writeExcel(excelRows, 1, u"电表")
                if meterStatus == "1":
                    writeExcel(excelRows, 2, u"离线", 2)
                elif meterStatus == "0":
                    writeExcel(excelRows, 2, u"在线")
                # 在 energy_device 中获取采集器 id
                collecterIDSql = "SELECT collectorId from energy_device where id=" + "\'" + str(elecmeterID) + "\'" + ";"
                collecterID=dbOperation(collecterIDSql)
                collecterID=collecterID[0][0]

                # 在采集器表中获取采集器的设备id
                collecterSQl="SELECT deviceId FROM collector_device WHERE id="+ "\'" + str(collecterID) + "\'" + ";"
                collectDeviceID=dbOperation(collecterSQl)
                collectDeviceID=collectDeviceID[0][0]
                writeExcel(excelRows, 3, collectDeviceID)

                # 获取设备表号
                shebeinmac = "SELECT address from energy_device where id=" + "\'" + str(elecmeterID) + "\'" + ";"
                shebeiMacResult=dbOperation(shebeinmac)
                writeExcel(excelRows, 4, shebeiMacResult[0][0])
                writeExcel(excelRows, 5,len(dbOperation(recoredCount)) )


                excelRows = excelRows + 1

                # # 操作日志系统
                # # "SELECT * FROM log_report WHERE device_id='41c0549225f9c71978e3f1c1c7f7637b' and url_path="status/report_water" ORDER BY res_time DESC; "
                # if meterType == "0":
                #     shuirizhiSql="SELECT report_content,report_time FROM log_report WHERE device_id=" + "\'" + str(collectDeviceID) + "\'" + "and url_path='status/report_water' AND report_content LIKE '%"+str(shebeiMacResult[0][0][2:])+"%' ORDER BY res_time DESC LIMIT 20 ";
                #     shuiriziResult = dbOperation(shuirizhiSql, db='danbay_task')
                #     if len(shuiriziResult)!=0:
                #         for shuirizhiIndex in shuiriziResult:
                #             # i[0]  payload, 第二个是时间
                #             shuirizhiPayload=shuirizhiIndex[0].split("payLoadString")[1]
                #             shuirizhiPayload=shuirizhiPayload.split("{")[1]
                #             shuirizhiShijian=shuirizhiIndex[1]
                #             if shebeiMacResult[0][0][2:] in shuirizhiPayload:
                #                 # 写日志表格
                #                 writeRiZhi(rizhiRow,0,homeaddress)
                #                 writeRiZhi(rizhiRow,1,u"水表")
                #
                #                 writeRiZhi(rizhiRow,2,collectDeviceID)
                #                 writeRiZhi(rizhiRow,3,shebeiMacResult[0][0])
                #                 # print rizhiPayload,rizhiShijian
                #                 writeRiZhi(rizhiRow,4,shuirizhiPayload)
                #                 writeRiZhi(rizhiRow,5,shuirizhiShijian,2)
                #                 rizhiRow=rizhiRow+1
                #     else:
                #         writeRiZhi(rizhiRow, 0, homeaddress)
                #         writeRiZhi(rizhiRow, 1, u"水表")
                #         writeRiZhi(rizhiRow, 2, collectDeviceID)
                #         writeRiZhi(rizhiRow, 3, shebeiMacResult[0][0])
                #         writeRiZhi(rizhiRow, 4, u'日志系统找不到该表对应的日志')
                #         rizhiRow = rizhiRow + 1
                #     print homeaddress,"完成对水表%s的日志检索" % shebeiMacResult[0][0]
                # elif meterType == "1":
                #     # rizhiSql = "SELECT report_content,report_time FROM log_report WHERE device_id=" + "\'" + str(
                #     #     collectDeviceID) + "\'" + "and url_path=\'status/report_kwh\' ORDER BY res_time DESC  "
                #
                #     rizhiSql="SELECT report_content,report_time FROM log_report WHERE device_id=" + "\'" + str(collectDeviceID) + "\'" + "and url_path=\'status/report_kwh\' AND report_content LIKE '%"+str(shebeiMacResult[0][0][2:])+"%' ORDER BY res_time DESC LIMIT 20 ";
                #     # print "电表"
                #     # print rizhiSql
                #     riziResult = dbOperation(rizhiSql, db='danbay_task')
                #     if len(riziResult)!=0:
                #         for rizhiIndex in riziResult:
                #             # i[0]  payload, 第二个是时间
                #             rizhiPayload=rizhiIndex[0].split("payLoadString")[1]
                #             rizhiPayload=rizhiPayload.split("{")[1]
                #             rizhiShijian=rizhiIndex[1]
                #             if shebeiMacResult[0][0] in rizhiPayload:
                #                 # 写日志表格
                #                 writeRiZhi(rizhiRow,0,homeaddress)
                #                 writeRiZhi(rizhiRow,1,u"电表")
                #                 writeRiZhi(rizhiRow,2,collectDeviceID)
                #                 writeRiZhi(rizhiRow,3,shebeiMacResult[0][0])
                #                 writeRiZhi(rizhiRow,4,rizhiPayload)
                #                 writeRiZhi(rizhiRow,5,rizhiShijian,2)
                #                 rizhiRow=rizhiRow+1
                #     else:
                #         writeRiZhi(rizhiRow, 0, homeaddress)
                #         writeRiZhi(rizhiRow, 1, u"电表")
                #         writeRiZhi(rizhiRow, 2, collectDeviceID)
                #         writeRiZhi(rizhiRow, 3, shebeiMacResult[0][0])
                #         writeRiZhi(rizhiRow, 4, u'日志系统找不到该表对应的日志')
                #         rizhiRow = rizhiRow + 1
                #
                #     print homeaddress,"完成对电表%s的日志检索"%shebeiMacResult[0][0]





    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表状态_' + unicode(a, "utf-8") + u'.xls')
    os.rename("rizhi.xls", unicode(address, "utf-8") + u'_日志系统_' + unicode(a, "utf-8") + u'.xls')
def getMeterReading(username,password,address):
    # excelRows记录当前Excel写到了多少行
    shuidianbiaoReadingRows = 1
    # generateCaiJiQi()
    generateShuiDianReading()
    # generateRiZhiBiao()
    # rizhiRow = 1
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount = getPageCount(username, password, address)

    # pageCount=46
    for index in range(pageCount):
        pageNo = index + 1
        payload = getPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        print "页面%s,响应记录是%s" % (str(pageNo), str(len(rsp)))
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)
            try:
                devicersp = json.loads(devicersp, encoding='utf-8')
            except:
                print "no json coulld be decode!!!!!!!!!!!!!!!"
                print devicersp
            devicersp = devicersp["result"]
            meterList = devicersp["meterList"]
            # 将水表或电表的状态写入表格
            for deviceIndex in range(len(meterList)):
                meter = meterList[deviceIndex]
                homeaddress = i["hosueInfo"]["homeAddress"]
                meterType = meter["meterType"]
                subType = meter["subType"]
                meterID = meter["deviceId"]
                elecmeterID = meter["id"]
                meterStatus = meter["onlineStatus"]
                aa = getHouseID(meterID)
                # "SELECT * FROM energy_day_consumption WHERE energyDevice=11510"
                # recoredCount="SELECT * FROM energy_day_consumption WHERE energyDevice=" + "\'" + str(elecmeterID) + "\'" + " ORDER BY readTime DESC LIMIT 11;"
                # recoredCount = "SELECT * FROM energy_day_consumption WHERE energyDevice=" + "\'" + str(
                #     elecmeterID) + "\'" + " ORDER BY readTime DESC;"
                writeExcel(shuidianbiaoReadingRows, 0, homeaddress)
                if meterType == "0":
                    if subType:
                        if subType == "1":
                            writeExcel(shuidianbiaoReadingRows, 1, u"热水")
                        elif subType == "0":
                            writeExcel(shuidianbiaoReadingRows, 1, u"冷水")
                elif meterType == "1":
                    writeExcel(shuidianbiaoReadingRows, 1, u"电表")
                if meterStatus == "1":
                    writeExcel(shuidianbiaoReadingRows, 2, u"离线", 2)
                elif meterStatus == "0":
                    writeExcel(shuidianbiaoReadingRows, 2, u"在线")

                # 在 energy_device 中获取采集器 id




                # # 在采集器表中获取采集器的设备id
                # collecterSQl = "SELECT deviceId FROM collector_device WHERE id=" + "\'" + str(collecterID) + "\'" + ";"
                # collectDeviceID = dbOperation(collecterSQl)
                # collectDeviceID = collectDeviceID[0][0]
                # writeExcel(excelRows, 3, collectDeviceID)

                # 获取设备表号
                shebeinmac = "SELECT address from energy_device where id=" + "\'" + str(elecmeterID) + "\'" + ";"
                shebeiMacResult = dbOperation(shebeinmac)

                shuidianReadingSql = "SELECT meterCount FROM energy_day_consumption WHERE energyDevice in (SELECT id from energy_device WHERE address='" +shebeiMacResult[0][0] +"' )  and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-03-22';"

                shuidianReadingResult=dbOperation(shuidianReadingSql)
                if shuidianReadingResult:

                    shuidianReading=shuidianReadingResult[0][0]
                    writeExcel(shuidianbiaoReadingRows, 4, shuidianReading)
                else:
                    writeExcel(shuidianbiaoReadingRows, 4,u'没有读数' )
                # print shuidianReadingResult
                shebeinmac = "SELECT address from energy_device where id=" + "\'" + str(elecmeterID) + "\'" + ";"
                shebeiMacResult = dbOperation(shebeinmac)
                writeExcel(shuidianbiaoReadingRows, 3, shebeiMacResult[0][0])



                shuidianbiaoReadingRows = shuidianbiaoReadingRows + 1
    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表表头读数_' + unicode(a, "utf-8") + u'.xls')


def LockOnlineStatus(username,password,address):
    excelRows = 1
    generateLockInfoExcel()
    payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pageCount = rsp["pageCount"]
    print "查询地址：" + address
    print "返回总页数：" + str(pageCount)
    # 为了调试方便，暂时设置为10
    # excelRows=1
    for index in range(pageCount):
        pageNo=index+1
        print "正在查询第%s页"%pageNo
        payload=getPayload(pageNo,address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        # print rsp
        # print len(rsp)
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            # print houseid
            devicersp=getJson(ck,houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            locklist=devicersp["lockList"]
            gatewalist=devicersp["gatewayList"]
            lockRoomName=devicersp["roomName"]
            # print lockRoomName
            for deviceIndex in range(len(locklist)):
                lock=locklist[deviceIndex]
                lcokDeviceId=lock["deviceId"]
                lcokOnlineStatus=lock["onlineStatus"]
                writeExcel(excelRows, 0, lockRoomName)
                if lcokOnlineStatus=="0":
                    writeExcel(excelRows,1 , u"在线")
                else:
                    writeExcel(excelRows, 1, u"离线",2)
                writeExcel(excelRows,2 , lcokDeviceId)
                # 门锁mac
                lockMacAddressSql="SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                lockMacAddressResult=dbOperation(lockMacAddressSql)
                lockMacAddress=lockMacAddressResult[0][0]
                writeExcel(excelRows, 3, lockMacAddress)

                # 从数据库中获取中控ID-----
                lockCenterControlIDSql = "SELECT centerControl from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                lockCenterControlID = dbOperation(lockCenterControlIDSql)
                # 循环中控，找到该门锁的id
                for gateway in gatewalist:
                    gwid = gateway["id"]
                    # 如果门锁的中控ID跟当前获取到的中控id一致，那么记录该中控在下状态
                    if gwid == lockCenterControlID[0][0]:

                        # 中控mac
                        # 中控在线状态
                        gwonlineStatus = gateway["onlineStatus"]
                        if gwonlineStatus == "1":
                        # 离线
                            writeExcel(excelRows, 4, u'离线',2)
                        elif gwonlineStatus == "0":
                            writeExcel(excelRows, 4, u'在线')
                        centerControlMacSql="SELECT deviceId,macAddress from center_control where id=" + "\'" + str(gwid) + "\'" + ";"
                        centerControlMacSqlResult=dbOperation(centerControlMacSql)
                        centerControlMac=centerControlMacSqlResult[0][1]
                        centerControlDevID=centerControlMacSqlResult[0][0]
                        print centerControlMac,centerControlDevID,gwid

                        writeExcel(excelRows, 5, centerControlDevID)
                        writeExcel(excelRows, 6, centerControlMac)


                    # 在线
                # 写中控id
                # print lockRoomName
                excelRows=excelRows+1


    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_门锁在线离线_' + unicode(a, "utf-8") + u'.xls')
def fensanLockOnlineStatus(username,password):
    #获取所有房源地址
    excelRows = 1
    generateLockInfoExcel()
    getalladdresssql="SELECT location FROM homesourcelist WHERE homeSourceProviderId=163"
    getalladdress=dbOperation(getalladdresssql,db='danbay_projects')
    for homeaddr in getalladdress:
        address=homeaddr[0]
        payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
        ck = getCookies(username, password)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        print "查询地址："+ address
        print "返回总页数："+str(pageCount)
        # 为了调试方便，暂时设置为10
        # excelRows=1
        for index in range(pageCount):
            pageNo=index+1
            print "正在查询第%s页"%pageNo
            payload=getPayload(pageNo,address)
            r = requests.post(host, data=payload, cookies=ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            rsp = rsp["resultList"]
            # print rsp
            # print len(rsp)
            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                # print houseid
                devicersp=getJson(ck,houseid)
                devicersp = json.loads(devicersp, encoding='utf-8')
                devicersp = devicersp["result"]
                locklist=devicersp["lockList"]
                gatewalist=devicersp["gatewayList"]
                lockRoomName=devicersp["roomName"]
                # print lockRoomName
                for deviceIndex in range(len(locklist)):
                    lock=locklist[deviceIndex]
                    lcokDeviceId=lock["deviceId"]
                    lcokOnlineStatus=lock["onlineStatus"]
                    writeExcel(excelRows, 0, lockRoomName)
                    if lcokOnlineStatus=="0":
                        writeExcel(excelRows,1 , u"在线")
                    else:
                        writeExcel(excelRows, 1, u"离线",2)
                    writeExcel(excelRows,2 , lcokDeviceId)
                    # 门锁mac
                    lockMacAddressSql="SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                    lockMacAddressResult=dbOperation(lockMacAddressSql)
                    lockMacAddress=lockMacAddressResult[0][0]
                    writeExcel(excelRows, 3, lockMacAddress)

                    # 从数据库中获取中控ID
                    lockCenterControlIDSql = "SELECT centerControl from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                    lockCenterControlID = dbOperation(lockCenterControlIDSql)
                    # 循环中控，找到该门锁的id
                    for gateway in gatewalist:
                        gwid = gateway["id"]
                        # 如果门锁的中控ID跟当前获取到的中控id一致，那么记录该中控在下状态
                        if gwid == lockCenterControlID[0][0]:
                            # 中控mac
                            # 中控在线状态
                            gwonlineStatus = gateway["onlineStatus"]
                            if gwonlineStatus == "1":
                            # 离线
                                writeExcel(excelRows, 4, u'离线',2)
                            elif gwonlineStatus == "0":
                                writeExcel(excelRows, 4, u'在线')
                            centerControlMacSql="SELECT deviceId,macAddress from center_control where id=" + "\'" + str(gwid) + "\'" + ";"
                            centerControlMacSqlResult=dbOperation(centerControlMacSql)
                            centerControlMac=centerControlMacSqlResult[0][1]
                            centerControlDevID=centerControlMacSqlResult[0][0]
                            print centerControlMac,centerControlDevID,gwid
                            writeExcel(excelRows, 5, centerControlDevID)
                            writeExcel(excelRows, 6, centerControlMac)
                        # 在线
                    # 写中控id
                    print lockRoomName
                    excelRows=excelRows+1


    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", u'分散式公寓' + u'_门锁在线离线_' + unicode(a, "utf-8") + u'.xls')
    # 碧桂园门锁预置密码

def BGYLockPwd(username, password):
    # 获取所有房源地址
    excelRows = 1
    generateLockPwdInfoExcel()
    getalladdresssql = "SELECT location FROM homesourcelist WHERE homeSourceProviderId=163"
    getalladdress = dbOperation(getalladdresssql, db='danbay_projects')
    for i in getalladdress:
        print i[0]
    for homeaddr in getalladdress:
        address = homeaddr[0]
        payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "",
                   "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
        ck = getCookies(username, password)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        print "查询地址："+ address
        print "返回总页数："+str(pageCount)
        # 为了调试方便，暂时设置为10
        # excelRows=1
        for index in range(pageCount):
            pageNo = index + 1
            print "正在查询第%s页" % pageNo
            payload = getPayload(pageNo, address)
            r = requests.post(host, data=payload, cookies=ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            rsp = rsp["resultList"]
            # print rsp
            # print len(rsp)
            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                # print houseid
                devicersp = getJson(ck, houseid)
                devicersp = json.loads(devicersp, encoding='utf-8')
                devicersp = devicersp["result"]
                locklist = devicersp["lockList"]
                gatewalist = devicersp["gatewayList"]
                lockRoomName = devicersp["roomName"]
                # print lockRoomName
                for deviceIndex in range(len(locklist)):
                    lock = locklist[deviceIndex]
                    lcokDeviceId = lock["deviceId"]
                    lcokOnlineStatus = lock["onlineStatus"]
                    writeExcel(excelRows, 0, lockRoomName)
                    writeExcel(excelRows, 1, lcokDeviceId)
                    #获取预置租客密码
                    getZuKePwd="SELECT * from lock_pre_password WHERE dev_id='523fd1b9a21226d418e0776d14a8abee' AND psw_type=3"
                    getZuKePwdSql="SELECT psw_value from lock_pre_password WHERE dev_id=" + "\'" + lcokDeviceId + "\'" + "AND psw_type=3 AND delete_state=0"+";"
                    zuKePwd=dbOperation(getZuKePwdSql)
                    try :
                        zkPwd=zuKePwd[0][0]
                        writeExcel(excelRows, 2, zkPwd)
                    except:
                        writeExcel(excelRows, 2, u'没有预置租客密码')

                    #获取预置临时密码1
                    getPreTempPwdSql = "SELECT psw_value from lock_pre_password WHERE dev_id=" + "\'" + lcokDeviceId + "\'" + "AND psw_type=0 AND delete_state=0" + ";"
                    PreTempPwd = dbOperation(getPreTempPwdSql)
                    # print  PreTempPwd
                    for ptpIndex in range (len(PreTempPwd)):
                        # print PreTempPwd[ptpIndex][0]
                        # print PreTempPwd[2+ptpIndex][0]
                        writeExcel(excelRows, 3+ptpIndex, PreTempPwd[ptpIndex][0])
                    #获取预置临时密码2
                    #获取预置临时密码3


                    print lockRoomName
                    excelRows = excelRows + 1

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", u'碧桂圆分散式公寓' + u'_门锁预置密码_' + unicode(a, "utf-8") + u'.xls')
def getCenterControlInfo(username, password,address):
    generateCenterControlinfo()
    addedCenterID = []
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount = getPageCount(username, password, address)
    # 获取中控id
    for index in range(pageCount):
        pageNo = index + 1
        payload = getPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            # # 通过 houseid 获取某个房间的设备
            req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
            r = requests.get(req, cookies=ck)
            # devicersp = r.text
            devicersp=getJson(ck, houseid)
            # devicersp = getJson(ck, houseid)

            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            gatewayList = devicersp["gatewayList"]


            for deviceIndex in range(len(gatewayList)):
                gw = gatewayList[deviceIndex]
                homeaddress = getHomeAddress(houseid)
                gwID = gw["id"]
                if gwID not in addedCenterID:
                    addedCenterID.append(gwID)
            print addedCenterID
    ck = getCookies(username, password)
    print "中控个数为：",len(addedCenterID)
    for controlIDIndex in range(len(addedCenterID)):
        devicersp = getCenterControlJson(ck,addedCenterID[controlIDIndex])
        devicersp = json.loads(devicersp, encoding='utf-8')
        devicersp = devicersp["result"]
        devicersp=devicersp["basicInfo"]
        centerContronDevmacAddress = devicersp["macAddress"]
        centerContronDevVersion = devicersp["deviceModel"]
        centerContronDevID=devicersp["deviceId"]
        centerContronDevaddress=devicersp["address"]
        writeExcel(controlIDIndex+1,0,centerContronDevmacAddress)
        writeExcel(controlIDIndex+1,1,centerContronDevVersion)
        writeExcel(controlIDIndex+1,2,centerContronDevID)
        writeExcel(controlIDIndex+1,3,centerContronDevaddress)

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')

    # 通过设备中心获取中控信息
def getCenterControlInfoByDeviceCenter(username, password,address):
    generateCenterControlInfoAtDeviceCenter()
    host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
    ck = getCookies(username, password)
    pageCount = centerControlGetPageCount(username, password, address)
    # 获取中控id
    excelRowCount=1
    for index in range(pageCount):

        pageNo = index + 1
        print "正在查询第%s页" % str(pageNo)
        payload = getCenterControlPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            resultrow = rsp[k]
            # ccid=centercontrolid
            ccid=resultrow["id"]
            # cc=centerControl
            ccMacSql="SELECT macAddress FROM center_control WHERE id="+ "\'" + str(ccid) + "\'" + ";"
            ccMac=dbOperation(ccMacSql)
            ccMac=ccMac[0][0]
            print ccMac
            houseAddress=resultrow["houseAddress"]
            centerControlstatus=resultrow["status"]
            centerCtrolVersion=resultrow["version"]

            writeExcel(excelRowCount, 0, ccMac)
            writeExcel(excelRowCount, 1, centerCtrolVersion)
            if centerControlstatus=="1":
                writeExcel(excelRowCount, 2, u"离线",2)
            else:
                writeExcel(excelRowCount, 2, u"在线")
            writeExcel(excelRowCount, 3, houseAddress)
            excelRowCount=excelRowCount+1
    #         # # 通过 houseid 获取某个房间的设备
    #         req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
    #         r = requests.get(req, cookies=ck)
    #         # devicersp = r.text
    #         devicersp = getJson(ck, houseid)
    #         # devicersp = getJson(ck, houseid)
    #
    #         devicersp = json.loads(devicersp, encoding='utf-8')
    #         devicersp = devicersp["result"]
    #         gatewayList = devicersp["gatewayList"]
    #
    #         for deviceIndex in range(len(gatewayList)):
    #             gw = gatewayList[deviceIndex]
    #             homeaddress = getHomeAddress(houseid)
    #             gwID = gw["id"]
    #             if gwID not in addedCenterID:
    #                 addedCenterID.append(gwID)
    #         print addedCenterID
    # ck = getCookies(username, password)
    # print "中控个数为：", len(addedCenterID)
    # for controlIDIndex in range(len(addedCenterID)):
    #     devicersp = getCenterControlJson(ck, addedCenterID[controlIDIndex])
    #     devicersp = json.loads(devicersp, encoding='utf-8')
    #     devicersp = devicersp["result"]
    #     devicersp = devicersp["basicInfo"]
    #     centerContronDevmacAddress = devicersp["macAddress"]
    #     centerContronDevVersion = devicersp["deviceModel"]
    #     centerContronDevID = devicersp["deviceId"]
    #     centerContronDevaddress = devicersp["address"]
    #     writeExcel(controlIDIndex + 1, 0, centerContronDevmacAddress)
    #     writeExcel(controlIDIndex + 1, 1, centerContronDevVersion)
    #     writeExcel(controlIDIndex + 1, 2, centerContronDevID)
    #     writeExcel(controlIDIndex + 1, 3, centerContronDevaddress)

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')



def getOfflineCenterControl(username, password,address):
    generateCenterControlExcel()
    addedCenterID = []
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount = getPageCount(username, password, address)
    for index in range(pageCount):
        pageNo=index+1
        payload=getPayload(pageNo,address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            # # 通过 houseid 获取某个房间的设备
            # req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
            # r = requests.get(req, cookies=ck)
            # devicersp = r.text
            devicersp=getJson(ck,houseid)

            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            gatewayList=devicersp["gatewayList"]
            for deviceIndex in range(len(gatewayList)):
                gw=gatewayList[deviceIndex]
                homeaddress=getHomeAddress(houseid)
                gwID=gw["id"]
                if gwID not in addedCenterID:
                    addedCenterID.append(gwID)
                    print addedCenterID
                    onlineDeviceNum=gw["onlineDeviceNum"]
                    onlineStatus=gw["onlineStatus"]
                    totalDeviceNum=gw["totalDeviceNum"]
                    roomAddr=devicersp["roomName"]
                    # print roomAddr
                    #房间地址
                    writeExcel(len(addedCenterID), 0, roomAddr)
                    # 中控id
                    writeExcel(len(addedCenterID), 1, gwID)
                    # 在线状态
                    if onlineStatus=="0":
                        writeExcel(len(addedCenterID), 2, u"在线")
                    elif onlineStatus=="1":
                        writeExcel(len(addedCenterID), 2, u"离线",2)

                    if onlineDeviceNum:
                        writeExcel(len(addedCenterID), 3, onlineDeviceNum)
                    if totalDeviceNum:
                        writeExcel(len(addedCenterID), 4, totalDeviceNum)
                    #中控版本，设备id，mac地址
                    getCenterControlmacDevIDVersion="SELECT deviceId,macAddress,version FROM center_control WHERE id=%s"%gwID
                    result=dbOperation(getCenterControlmacDevIDVersion)

                    writeExcel(len(addedCenterID) ,5 , result[0][0])
                    writeExcel(len(addedCenterID) ,6 , result[0][1])
                    writeExcel(len(addedCenterID) ,7, result[0][2])
                # houseID，如果不为空就写进去

                #index*10+k*2+1+deviceIndex 行数    4 列数， meterID 需要写入的值
                # writeExcel(index*20+k*2+1+deviceIndex,4,meterID)

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_中控在线离线状态_' + unicode(a, "utf-8") + u'.xls')
def getConf():
    cwd=os.getcwd()
    with open(cwd+"\info.txt", 'r') as f:
        accountInfo=f.readlines()

    address=accountInfo[2].strip().split("=")[1]
    address=address.split(",")
    # print address
    account={}
    account["username"]=accountInfo[0].strip().split("=")[1]
    account["password"]=accountInfo[1].strip().split("=")[1]
    account["address"]=address

    account["option"] = accountInfo[4].strip().split("=")[1]

    return account
def getMeterCount(username, password, address):
    meterExcel()
    waterExcel()
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount = getPageCount(username, password, address)
    meterRows=1
    waterRows=1
    for index in range(pageCount):
        pageNo = index + 1
        print "正在获取第%s页"%pageNo
        payload = getPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)

            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            meterList = devicersp["meterList"]
            for deviceIndex in range(len(meterList)):
                meter = meterList[deviceIndex]
                homeaddress = i["hosueInfo"]["homeAddress"]
                meterType = meter["meterType"]
                subType = meter["subType"]
                meterID = meter["deviceId"]
                elecmeterID = meter["id"]
                meterStatus = meter["onlineStatus"]
                aa = getHouseID(meterID)
                # 0 水表
                mydatelist=[21,20,19,18,17,16,15]
                myseconddatelist=[20,21,22,23,24,25,26]
                if meterType == "0":
                    sql1 = "SELECT id,houseId from energy_device WHERE deviceId='" + meterID + "';"
                    addrAndId = dbOperation(sql1)
                    # roomaddress = addrAndId[0][1]
                    roomaddress = homeaddress
                    waterID = addrAndId[0][0]
                    # sql = "SELECT * FROM energy_day_consumption WHERE energyDevice='" + str(
                    #     waterID) + "' ORDER BY readTime DESC;"

                    # 1月21号的值
                    for mydate in mydatelist:
                        w1sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
                        waterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-01-"+str(mydate)+"';"
                        firstwaterCount = dbOperation(w1sql)
                        if firstwaterCount:
                            fwaterCount= firstwaterCount[0][0]
                            print w1sql,"水表1"
                            break
                    # 2月22或以后的值
                    for myseconddate in myseconddatelist:
                        w2sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
                        waterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-02-"+str(myseconddate)+"';"
                        secondwaterCount = dbOperation(w2sql)
                        if secondwaterCount:
                            SeCount= secondwaterCount[0][0]
                            print w2sql,"水表2"
                            break
                    StotalWater=SeCount-fwaterCount
                    print roomaddress,  StotalWater
                    TwowaterWriteExcel(waterRows, 0, roomaddress)
                    TwowaterWriteExcel(waterRows, 1, StotalWater)
                    TwowaterWriteExcel(waterRows, 2, waterID)
                    waterRows=waterRows+1

                elif meterType == "1":
                    metersql1 = "SELECT id,houseId from energy_device WHERE deviceId='" + meterID + "';"
                    meteraddrAndId = dbOperation(metersql1)
                    # meterroomaddress = meteraddrAndId[0][1]
                    meterID = meteraddrAndId[0][0]

                    # 1月21号的值
                    for myddate in mydatelist:
                        d1sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
                            meterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-01-"+str(myddate)+"';"
                        # print  sql
                        firstDianCount = dbOperation(d1sql)
                        if firstDianCount:
                            fDianCount= firstDianCount[0][0]
                            print d1sql,"电表1"
                            break

                    # 2月22或以后的值
                    for mysecondddate in myseconddatelist:
                        d2sql = "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
                            meterID) + "' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-02-" + str(
                            mysecondddate) + "';"
                        seconddianCount = dbOperation(d2sql)
                        if seconddianCount:
                            SedianCount = seconddianCount[0][0]
                            print d2sql,"电表2"
                            break

                    DtotalDian=SedianCount-fDianCount
                    TwometerWriteExcel(meterRows, 0, homeaddress,2)
                    TwometerWriteExcel(meterRows,1,DtotalDian,2)
                    TwometerWriteExcel(meterRows,2,meterID,2)
                    # meterResult = dbOperation(metersql)
                    # # print meterResult
                    # if meterResult:
                    #     meterCount = meterResult[0][2]
                    #     meterreadTime = meterResult[0][4]
                    #     # print meterroomaddress, meterreadTime, meterCount
                    #     # meterWriteExcel()
                    #     meterWriteExcel(meterRows,0,meterroomaddress)
                    #     meterWriteExcel(meterRows,1,meterreadTime,2)
                    #     meterWriteExcel(meterRows,2,meterCount)
                    # else:
                    #     meterWriteExcel(meterRows,0,meterroomaddress+"-"+metersql)

                    meterRows=meterRows+1

def getAddress(username, password):
    addrSql="SELECT detailAddress FROM homesourcelist WHERE deleteState !=1 AND location LIKE '%深圳%' ;"
    addrResult=dbOperation(addrSql,db='danbay_projects')
    excelRowCount = 1
    generateCenterControlInfoAtDeviceCenter()

    # for dbaddr in addrResult:
    #     if dbaddr[0]:
    #         print dbaddr[0]
    #         host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
    #         ck = getCookies(username, password)
    #         # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
    #         pageCount = centerControlGetPageCount(username, password, dbaddr[0])
    #         # 获取中控id
    #         if pageCount !=0:
    #             for index in range(pageCount):
    #                 pageNo = index + 1
    #                 print "正在查询第%s页" % str(pageNo)
    #                 payload = getCenterControlPayload(pageNo, dbaddr[0])
    #                 r = requests.post(host, data=payload, cookies=ck)
    #                 rsp = json.loads(r.text)
    #                 rsp = rsp["result"]["result"]
    #                 rsp = rsp["resultList"]
    #                 for k in range(len(rsp)):
    #                     resultrow = rsp[k]
    #                     ccid = resultrow["id"]
    #                     # cc=centerControl
    #                     ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
    #                     ccMac = dbOperation(ccMacSql)
    #                     ccMac = ccMac[0][0]
    #                     print ccMac
    #                     houseAddress = resultrow["houseAddress"]
    #                     centerControlstatus = resultrow["status"]
    #                     centerCtrolVersion = resultrow["version"]
    #                     print houseAddress,centerControlstatus,centerCtrolVersion
    #
    #                     writeExcel(excelRowCount, 0, ccMac)
    #                     writeExcel(excelRowCount, 1, centerCtrolVersion)
    #                     if centerControlstatus == "1":
    #                         writeExcel(excelRowCount, 2, u"离线", 2)
    #                     else:
    #                         writeExcel(excelRowCount, 2, u"在线")
    #                     writeExcel(excelRowCount, 3, houseAddress)
    #                     excelRowCount = excelRowCount + 1
    #                     print excelRowCount, "Excel 行数"



    # 获取所有项目的中控在线离线状态
    excelRowCount = 1
    generateCenterControlInfoAtDeviceCenter()
    host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
    ck = getCookies(username, password)
    # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
    # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
    payload = getCenterControlPayload("1", "深圳")
    host = "http://www.danbay.cn/system/centerControl/findCenterControlSetting"
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]["result"]
    pageCount = rsp["pageCount"]

    print "总共的页数是:", pageCount
    # return pageCount
    # 获取中控id
    if pageCount != 0:
        for index in range(pageCount):
            pageNo = index + 1
            print "正在查询第%s页" % str(pageNo)
            # payload = getCenterControlPayload(pageNo, dbaddr[0])
            payload = {'pageNo': pageNo, 'pageSize': '8', "status": "", 'likeStr': "深圳", "isNewVersion": ""}
            # return payload
            r = requests.post(host, data=payload, cookies=ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]["result"]
            rsp = rsp["resultList"]
            for k in range(len(rsp)):
                resultrow = rsp[k]
                ccid = resultrow["id"]
                # cc=centerControl
                ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
                ccMac = dbOperation(ccMacSql)
                ccMac = ccMac[0][0]
                print ccMac
                houseAddress = resultrow["houseAddress"]
                centerControlstatus = resultrow["status"]
                centerCtrolVersion = resultrow["version"]
                print houseAddress, centerControlstatus, centerCtrolVersion

                writeExcel(excelRowCount, 0, ccMac)
                writeExcel(excelRowCount, 1, centerCtrolVersion)
                if centerControlstatus == "1":
                    writeExcel(excelRowCount, 2, u"离线", 2)
                else:
                    writeExcel(excelRowCount, 2, u"在线")
                writeExcel(excelRowCount, 3, houseAddress)
                excelRowCount = excelRowCount + 1
                print excelRowCount, "Excel 行数"
    # print  "done"
    # a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    # os.rename("test.xls",  u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')

if __name__ == '__main__':
    starttime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    start = datetime.datetime.now()
    print "开始计时: %s" % starttime
    # print "Please wait, Getting the numbers of lock user....."
    config = getConf()
    #######################################################
    #集中式公寓
    # read_excel()
    print "地址为：",config["address"][0]
    print "选项是：",config["option"]

    if  config["option"]==str(1):
     getDeviceListAllWithCorrectInfo(config["username"],config["password"],config["address"]) # 获取门锁预置密码以及正式密码的总数
    elif  config["option"]==str(2):
        getAmmeterDeviceId(config["username"],config["password"],config["address"]) #获取水电表在线离线状态 以及 对应中控在线离线状态
    elif config["option"] == 3:

        getCountCaijiqi(config["username"],config["password"],config["address"]) #获取采集器下的水电表在线离线状态  以及采集器的设备id 和水电表表号
    elif config["option"] == 4:
        getOfflineCenterControl(config["username"],config["password"],config["address"]) #获取中控的设备id， mac 地址，在线离线状态
    # elif config["option"] == 5:
    #     getCenterControlInfo(config["username"],config["password"],config["address"])  # 先爬去所有中控的id，然后再根据id 爬取中控的在线离线已经设备信息
    elif config["option"] == 5:
        getCenterControlInfoByDeviceCenter(config["username"],config["password"],config["address"]) #在设备中心获取中控的信息在线离线，以及 mac信息
    elif config["option"] == 6:
        LockOnlineStatus(config["username"],config["password"],config["address"]) #查看门锁在线离线状态
    # elif config["option"] == 7:
    #     getMeterCount(config["username"],config["password"],config["address"]) #厦门青年公寓计算2月份月冻结数据
    elif config["option"] == 7:
        getMeterReading(config["username"],config["password"],config["address"]) #获取水电表表头读数

        # getAddress(config["username"],config["password"])

    # 电表部分
    # 集中式公寓
    ############################################################################################

    ##############################################################################################
    # 分散式公寓
    # fensanLockOnlineStatus(config["username"],config["password"])
    # BGYLockPwd(config["username"],config["password"])





    ##############################################################################################
    endtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    end = datetime.datetime.now()
    print "结束计时: %s" % endtime
    print "总共花费时间为: %s" % (end - start)
    # raw_input("Press any key to exit....")