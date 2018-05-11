#!usr/bin/env python
#-*- coding:utf-8 _*-
"""
@author:albert.chen
@file: excelOperate.py
@time: 2018/05/11/9:04
"""
import xlrd
import xlwt
from xlutils.copy import copy
from xlwt import *
class ExcelTool():
    @staticmethod
    def set_style(name, height, bold=False):
        style = xlwt.XFStyle()  # 初始化样式

        font = xlwt.Font()  # 为样式创建字体
        font.name = name  # 'Times New Roman'
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        return style
    @staticmethod
    def redStyle():
        red_style = xlwt.XFStyle()  # 初始化样式
        pattern = Pattern()  # 创建一个模式
        pattern.pattern = Pattern.SOLID_PATTERN  # 设置其模式为实型
        pattern.pattern_fore_colour = 2
        red_style.pattern = pattern
        return red_style

    @staticmethod
    def dateStyle():
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'yyy/mm/dd'
        return date_format
    @staticmethod
    def generateExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        '''
        创建第一个sheet:
            sheet1
        '''
        sheet1 = f.add_sheet(u'门锁密码统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'门锁管家密码数', u'门锁租客密码数', u'门锁临时密码数', u'云端租客密码数', u'云端管家密码数', u'门锁已使用密码数', u'门锁密码容量']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 50
            sheet1.col(1).width = 256 * 18
            sheet1.col(2).width = 256 * 28
            sheet1.col(3).width = 256 * 28
            sheet1.col(4).width = 256 * 18
            sheet1.col(5).width = 256 * 18
            sheet1.col(6).width = 256 * 15
            sheet1.col(7).width = 256 * 15
            # sheet1.col(7).width = 2+56 * 50
            # sheet1.col(8).width = 256 * 15
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def meterExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'电表电量记录', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间号', u'冻结时间', u'电表读数']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 30
            sheet1.col(1).width = 256 * 18
            sheet1.col(2).width = 256 * 15
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def waterExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'水表水量记录', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间号', u'用量']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 30
            sheet1.col(1).width = 256 * 18
            sheet1.col(2).width = 256 * 15
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("water.xls")
    @staticmethod
    def generateLockInfoExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        '''
        创建第一个sheet:
            sheet1
        '''
        sheet1 = f.add_sheet(u'门锁在线离线状态', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'房间地址', u'门锁状态', u'门锁设备ID', u'门锁Mac', u'中控状态', u'中控设备ID', u'中控Mac', u'中控版本']
        # 生成第一行,并设置单元格长度，只需设置一次即可状态
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 60
            sheet1.col(1).width = 256 * 10
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 20
            sheet1.col(4).width = 256 * 10
            sheet1.col(5).width = 256 * 50
            sheet1.col(6).width = 256 * 20
            sheet1.col(6).width = 256 * 20
            # sheet1.col(7).width = 2 + 56 * 50
            # sheet1.col(8).width = 256 * 15
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateLockPwdInfoExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        '''
        创建第一个sheet:
            sheet1
        '''
        sheet1 = f.add_sheet(u'碧桂园门锁预置密码表', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'设备ID', u'预置租客密码', u'预置临时密码1', u'预置临时密码2', u'预置临时密码3']
        # 生成第一行,并设置单元格长度，只需设置一次即可状态
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 60
            sheet1.col(1).width = 256 * 20
            sheet1.col(1).width = 256 * 15
            sheet1.col(2).width = 256 * 16
            sheet1.col(3).width = 256 * 16
            sheet1.col(4).width = 256 * 16
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateMeterExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'水电表信息统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'类型', u'状态', u'30天内离线总时长(单位:天)', u'30天内离线次数', u'当前读数', u'设备Mac', u'设备ID', u'中控mac', u'中控状态',
                u'中控device id']
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 40
            sheet1.col(1).width = 256 * 6
            sheet1.col(2).width = 256 * 6
            sheet1.col(3).width = 256 * 25
            sheet1.col(4).width = 256 * 14
            sheet1.col(5).width = 256 * 11
            sheet1.col(6).width = 256 * 18
            sheet1.col(7).width = 256 * 33
            sheet1.col(8).width = 256 * 16
            sheet1.col(9).width = 256 * 12
            sheet1.col(10).width = 256 * 33
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def dushuExcel():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'水电表信息统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'设备类型', u'设备状态', u'设备框读数', u'当前读数']
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 50
            sheet1.col(1).width = 256 * 12
            sheet1.col(2).width = 256 * 12
            sheet1.col(3).width = 256 * 12
            sheet1.col(4).width = 256 * 40
            sheet1.col(5).width = 256 * 20
            sheet1.col(6).width = 256 * 20
            sheet1.col(7).width = 256 * 12
            sheet1.col(8).width = 256 * 50
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateShuiDianExcelFromDB():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'水电表信息统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'类型', u'状态', u'设备Mac', u'设备ID', u'表头读数', u'离线时间', u'在线时间', u'中控状态', u'中控id', u'中控mac',
                u'中控版本号',
                u'中控关联地址']
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 50  # 房间地址
            sheet1.col(1).width = 256 * 8  # 类型
            sheet1.col(2).width = 256 * 8  # 状态
            sheet1.col(3).width = 256 * 12  # 设备Mac
            sheet1.col(4).width = 256 * 35  # 设备ID
            sheet1.col(5).width = 256 * 11  # 表头读数
            sheet1.col(6).width = 256 * 20  # 离线时间
            sheet1.col(7).width = 256 * 20  # 在线时间
            sheet1.col(8).width = 256 * 10  # 中控状态
            sheet1.col(9).width = 256 * 35  # 中控id
            sheet1.col(10).width = 256 * 20  # 中控mac
            sheet1.col(11).width = 256 * 15  # 中控版本号
            sheet1.col(12).width = 256 * 50  # 中控关联地址
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateCaiJiQi():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        sheet1 = f.add_sheet(u'水电表信息统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'类型', u'状态', u'30天内离线总时长(单位:天)', u'30天内离线次数', u'当前读数', u'采集器Mac', u'采集器ID', u'设备表号', u'设备ID',
                u'中控Mac', u'中控状态', u'中控deviceId'
                ]
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 40
            sheet1.col(1).width = 256 * 8
            sheet1.col(2).width = 256 * 8
            sheet1.col(3).width = 256 * 25
            sheet1.col(4).width = 256 * 14
            sheet1.col(5).width = 256 * 13
            sheet1.col(6).width = 256 * 18
            sheet1.col(7).width = 256 * 20
            sheet1.col(8).width = 256 * 12
            sheet1.col(9).width = 256 * 20
            sheet1.col(10).width = 256 * 18
            sheet1.col(11).width = 256 * 13
            sheet1.col(12).width = 256 * 27
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateShuiDianReading():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        '''
        创建第一个sheet:
            sheet1
        '''
        sheet1 = f.add_sheet(u'水电表表头读数', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'设备类型', u'在线状态', u'设备表号', u'设备当前读数', u'读取时间', u"采集器ID"]
        # 生成第一行,并设置单元格长度，只需设置一次即可
        # date_format = xlwt.XFStyle()
        # date_format.num_format_str = 'yyyy-mm-dd'
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 50
            sheet1.col(1).width = 256 * 12
            sheet1.col(2).width = 256 * 15
            sheet1.col(3).width = 256 * 20
            sheet1.col(4).width = 256 * 15
            sheet1.col(5).width = 256 * 20
            sheet1.col(6).width = 256 * 50

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))

        f.save("test.xls")
    @staticmethod
    def generateRiZhiBiao():
        f = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        '''
        创建第一个sheet:
            sheet1
        '''
        sheet1 = f.add_sheet(u'日志表信息', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'房间地址', u'设备类型', u'采集器DevID', u'设备表号', u'PayLoad', u'上报时间']
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

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("rizhi.xls")
    @staticmethod
    def generateCenterControlExcel():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'房间地址', u'中控id', u'中控状态', u'该中控下的在线设备数', u'该中控设备总数', u'中控deviceID', u'中控Mac地址', u'中控版本']
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
            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateCenterControlinfo():
        # centerContronDevmacAddress = devicersp["macAddress"]
        # centerContronDevVersion = devicersp["deviceModel"]
        # centerContronDevID = devicersp["deviceId"]
        # centerContronDevaddress = devicersp["address"]
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'中控MAC地址', u'中控版本', u'中控设备ID', u'中控地址', u'中控在线离线状态']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 20
            sheet1.col(1).width = 256 * 23
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 150

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def read_excel():
        # 文件位置

        ExcelFile = xlrd.open_workbook(u'安徽省合肥市蜀山区稻香路与山湖路交口向西100米_水电表状态_2018-02-27_23_29_07.xls')

        # 获取目标EXCEL文件sheet名

        print ExcelFile.sheet_names()[0]

        # ------------------------------------

        # 若有多个sheet，则需要指定读取目标sheet例如读取sheet2

        # sheet2_name=ExcelFile.sheet_names()[1]

        # ------------------------------------

        # 获取sheet内容【1.根据sheet索引2.根据sheet名称】

        sheet = ExcelFile.sheet_by_index(0)

        # sheet=ExcelFile.sheet_by_name('TestCase002')

        # 打印sheet的名称，行数，列数

        print sheet.name, sheet.nrows, sheet.ncols
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
        print sheet.cell(1, 0).value.encode('utf-8')
        #
        # print sheet.cell_value(1,0).encode('utf-8')
        #
        # print sheet.row(1)[0].value.encode('utf-8')

        # 打印单元格内容格式
    @staticmethod
    def generateCenterControlInfoAtDeviceCenter():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'中控MAC地址', u'中控版本', u'中控在线离线状态', u'中控地址']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 20
            sheet1.col(1).width = 256 * 23
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 150

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateCheckLockSync():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'门锁密码同步次数统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'门锁地址', u'门锁Mac', u'门锁ID', u'日期：同步次数']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 80
            sheet1.col(1).width = 256 * 23
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 200

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateDeviceOffline():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'设备离线时间统计', cell_overwrite_ok=True)  # 创建sheet
        row0 = [u'门锁地址', u'门锁Mac', u'门锁ID', u'离线总时长(单位：天)']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 80
            sheet1.col(1).width = 256 * 23
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 15

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateLockPwdCount():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'门锁密码记录数统计', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'房间地址', u'门锁租客密码记录数', u'设备ID']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 80
            sheet1.col(1).width = 256 * 20
            sheet1.col(2).width = 256 * 40
            # sheet1.col(3).width = 256 * 150

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def generateAllCenterControlInfoFromDataBase():
        f = xlwt.Workbook(encoding='utf-8')
        sheet1 = f.add_sheet(u'中控信息统计', cell_overwrite_ok=True)  # 创建sheet
        # row0 = [u'房间地址',u'管理员密码个数',u'管家密码个数',u'租客密码个数',u'临时密码个数',u'预置租客个数',u'预置临时个数',u'密码已使用数',u'门锁ID']
        row0 = [u'中控MAC地址', u'中控版本', u'中控在线离线状态', u'中控地址']
        # 生成第一行,并设置单元格长度，只需设置一次即可
        for i in range(0, len(row0)):
            sheet1.col(0).width = 256 * 20
            sheet1.col(1).width = 256 * 23
            sheet1.col(2).width = 256 * 40
            sheet1.col(3).width = 256 * 150

            sheet1.write(0, i, row0[i], ExcelTool.set_style('Times New Roman', 220, True))
        f.save("test.xls")
    @staticmethod
    def writeExcel(rowIndex, colIndex, cellValue, userStyle=1):
        '''
        该函数主要功能是实现往已存在的Excel表格中写入数据
        :param workBookNmae: Excel 表名
        :param rowIndex: 行索引
        :param colIndex: 列索引
        :param cellValue: 单元格的值
        :return:
        '''

        rb = xlrd.open_workbook("test.xls", formatting_info=True)
        # sheet=data.sheet_by_name(u'门锁密码统计')
        wb = copy(rb)
        # write(1, 0, "test")，第一个是行索引，第二个是列索引，第三个是单元格的值
        if userStyle == 1:
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue)

        elif userStyle == 3:
            date_format = xlwt.XFStyle()
            date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
        else:
            userStyle = ExcelTool.redStyle()
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, userStyle)
        wb.save(u'test.xls')
    @staticmethod
    def writeRiZhi(rowIndex, colIndex, cellValue, userStyle=1):
        '''
        该函数主要功能是实现往已存在的Excel表格中写入数据
        :param workBookNmae: Excel 表名
        :param rowIndex: 行索引
        :param colIndex: 列索引
        :param cellValue: 单元格的值
        :return:
        '''

        rb = xlrd.open_workbook("rizhi.xls", formatting_info=True)
        # sheet=data.sheet_by_name(u'门锁密码统计')

        # write(1, 0, "test")，第一个是行索引，第二个是列索引，第三个是单元格的值
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
        # rb = xlrd.open_workbook("water.xls", formatting_info=True)
        wb = copy(rb)
        if userStyle == 1:
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
        else:
            userStyle = ExcelTool.redStyle()
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
        wb.save(u'rizhi.xls')
    @staticmethod
    def waterWriteExcel(rowIndex, colIndex, cellValue, userStyle=1):
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'yyyy/mm/dd'
        rb = xlrd.open_workbook("water.xls", formatting_info=True)
        wb = copy(rb)
        if userStyle == 1:
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
        else:
            userStyle = ExcelTool.redStyle()
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
        wb.save(u'water.xls')
    @staticmethod
    def TwowaterWriteExcel(rowIndex, colIndex, cellValue, userStyle=1):
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
    @staticmethod
    def meterWriteExcel(rowIndex, colIndex, cellValue, userStyle=1):
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'yyyy/mm/dd'
        rb = xlrd.open_workbook("test.xls", formatting_info=True)
        wb = copy(rb)
        if userStyle == 1:
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
        else:
            userStyle = ExcelTool.redStyle()
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
        wb.save(u'test.xls')
    @staticmethod
    def TwometerWriteExcel(rowIndex, colIndex, cellValue, userStyle=1):
        date_format = xlwt.XFStyle()
        # date_format.num_format_str = 'yyyy/mm/dd'
        rb = xlrd.open_workbook("test.xls", formatting_info=True)
        wb = copy(rb)
        if userStyle == 1:
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue)
        else:
            userStyle = ExcelTool.redStyle()
            wb.get_sheet(0).write(rowIndex, colIndex, cellValue, date_format)
        wb.save(u'test.xls')
