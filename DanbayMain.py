#!usr/bin/env python
# -*- coding:utf-8 _*-
"""
@author:albert.chen
@file: DanbayMain.py
@time: 2018/03/25/20:44
"""

import random
import sys
import threading
import time

import MySQLdb
import datetime
import json
import os
import requests
import xlrd
import xlwt
from xlutils.copy import copy
from xlwt import *

from Util.excelOperate import ExcelTool as et
defaultencoding = 'utf-8'
if sys.getdefaultencoding() != defaultencoding:
    reload(sys)
    sys.setdefaultencoding(defaultencoding)

setSleepTime=50
setSleepTimeInfo=u"开始休眠%s秒"%setSleepTime

def renameFile(fileName):
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(fileName, "utf-8") + u'_门锁密码容量_' + unicode(a, "utf-8") + u'.xls')

def dbOperation(sql, db='danbay_device'):
    conn = MySQLdb.connect(
        host='rm-wz916f30z77a773rdo.mysql.rds.aliyuncs.com',
        port=3306,
        user='XXX',
        passwd='XXXXXXXXX@',
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
    sql = "SELECT psw_alias,psw_type from lock_pre_password WHERE dev_id=" + "\'" + deviceID + "\'" + "and delete_state !=1;"
    result = dbOperation(sql)
    renter_pwd = 0
    tmp_pwd = 0
    for row in result:
        if row[1] == "3":
            renter_pwd = renter_pwd + 1
        if row[1] == "0":
            tmp_pwd = tmp_pwd + 1
    # print renter_pwd,tmp_pwd
    pre_pwd_count = {}
    pre_pwd_count["renter_pwd"] = renter_pwd
    pre_pwd_count["tmp_pwd"] = tmp_pwd
    return pre_pwd_count


def getPwdCountsInNormal(deviceID):
    '''
    获取正常密码表的密码数量
    :param sql: 需要执行的sql语句
    :return: 返回密码总数的字典
    '''
    device_info_sql = "SELECT id FROM device_info WHERE deviceId=" + "\'" + deviceID + "\'" + ";"
    result = dbOperation(device_info_sql)
    id = 0
    for row in result:
        id = row[0]
    device_pwd_info_sql = "SELECT pwdType,pwdAlias from device_pwd_info WHERE deviceInfo=" + "\'" + str(id) + "\'" + ";"
    pwdresult = dbOperation(device_pwd_info_sql)
    housekeeper = 0
    admin = 0
    renter = 0
    tmp = 0
    for pwdinfo in pwdresult:
        if pwdinfo[0] == "0":
            tmp = tmp + 1
        elif pwdinfo[0] == "1":
            admin = admin + 1
        elif pwdinfo[0] == "2":
            housekeeper = housekeeper + 1
        elif pwdinfo[0] == "3":
            renter = renter + 1
    # print tmp,admin,housekeeper,renter
    nor_pwd_count = {}
    nor_pwd_count["tmp"] = tmp
    nor_pwd_count["admin"] = admin
    nor_pwd_count["housekeeper"] = housekeeper
    nor_pwd_count["renter"] = renter
    # print nor_pwd_count
    return nor_pwd_count


def getHomeAddress(ID):
    # 'SELECT homeAddress FROM house_info WHERE ID='24052''
    device_info_sql = "SELECT homeAddress FROM house_info WHERE ID=" + "\'" + str(ID) + "\'" + ";"
    result = dbOperation(device_info_sql)
    homeaddress = result[0]
    return homeaddress[0]
    # print  homeaddress[0]


def getHouseID(devuiceID):
    # 'SELECT * from energy_device where deviceId='a1bb95ed70878d2d729b204e140eff4e''
    houseIDSql = "SELECT houseId from energy_device where deviceId=" + "\'" + str(devuiceID) + "\'" + ";"
    result = dbOperation(houseIDSql)
    houseID = result[0]
    # print houseID[0]
    return houseID[0]


def getCookies(username, password):
    loginurl = 'http://www.danbay.cn/system/goLoginning'
    payload = {'mc_username': username, 'mc_password': password, 'rememberMe': ""}
    r = requests.post(loginurl, data=payload)
    return r.cookies


def getPayload(pageNo, address):
    payload = {'pageNo': pageNo, 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "",
               "spaceType": "",
               "userId": "1"}
    return payload


def getCenterControlPayload(pageNo, address):
    payload = {'pageNo': pageNo, 'pageSize': '8', "status": "", 'likeStr': address, "isNewVersion": ""}
    return payload


def getJson(ck, houseid):
    req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
    for i in range(15):
        try:
            r = requests.get(req, cookies=ck)
            devicersp = r.text
            if r.status_code == 200:
                if u"服务器出错了" in devicersp:
                    print u"通过HouseId 获取设备信息失败，尝试次数为：%s" % i
                    print (u"{occurTime}第{tryTime}次获取json数据,响应结果为：{rsp}".format(
                        occurTime=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())), tryTime=str(i),
                        rsp=devicersp))
                    time.sleep(setSleepTime)
                else:
                    return devicersp
            else:
                print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code,
                                                                           rspcontent=r.text)
        except:
            print devicersp.decode("utf-8")
            print req.decode("utf-8")
            time.sleep(setSleepTime)

def getCommonJson(host, payload, ck, reqMethod):
    for i in range(15):
        if reqMethod == "post":
            r = requests.post(host, data=payload, cookies=ck)
        else:
            r = requests.get(host, cookies=ck)
        devicersp = r.text

        if r.status_code == 200:

            if u"服务器出错了" in devicersp:
                print u"通过HouseId 获取设备信息失败，尝试次数为：%s" % i
                print (u"{occurTime}第{tryTime}次获取json数据,响应结果为：{rsp}".format(
                    occurTime=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())), tryTime=str(i),
                    rsp=devicersp))
                time.sleep(setSleepTime)
            else:
                return devicersp
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host.decode("utf-8")
            setSleepTimeInfo
            time.sleep(setSleepTime)
def getCenterControlJson(ck, centerControlID):
    for i in range(10):
        centerControlURL = "http://www.danbay.cn/system/centerControl/getCenterControlDetail?id=%s" % centerControlID
        r = requests.post(centerControlURL, cookies=ck)
        devicersp = r.text
        if r.status_code == 200:
            # //由于devicersp 是unicode 类型，所以要转一下if u"服务器出错了" in devicersp:
            if u"服务器出错了" in devicersp:
                print u"通过中控ID 获取设备信息失败，尝试次数为：%s" % i
                # 如果服务器报错了，休眠1分钟，然后再请求
                setSleepTimeInfo
                time.sleep(setSleepTime)
            else:
                return devicersp

        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print centerControlURL
            setSleepTimeInfo
            time.sleep(setSleepTime)
def getPageCount(username, password, address):
    payload = getPayload("1", address)
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    for tryTims in range(10):

        r = requests.post(host, data=payload, cookies=ck)

        if r.status_code==200:
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            pageCount = rsp["pageCount"]
            print "总共的页数是:".decode("utf-8"), pageCount
            return pageCount
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host.decode("utf-8")
            setSleepTimeInfo
            time.sleep(setSleepTime)

def getResultList(host, pageNo, address, ck):
    payload = getPayload(pageNo, address)
    for tryTime in range(10):

        r = requests.post(host, data=payload, cookies=ck)
        if r.status_code==200:
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            rsp = rsp["resultList"]
            return rsp
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host.decode("utf-8")
            setSleepTimeInfo
            time.sleep(setSleepTime)
def centerControlGetPageCount(username, password, address,fensan=1):
    payload = getCenterControlPayload("1", address)
    host = "http://www.danbay.cn/system/centerControl/findCenterControlSetting"
    ck = getCookies(username, password)

    for tryTims in range(10):
        r = requests.post(host, data=payload, cookies=ck)
        if r.status_code == 200:
            rsp = json.loads(r.text)
            rsp = rsp["result"]["result"]
            pageCount = rsp["pageCount"]
            if fensan !=1:
                print u"总共的页数是:%s"%pageCount
            return pageCount
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host.decode("utf-8")
            setSleepTimeInfo
            time.sleep(setSleepTime)
def getCenterControlpageCount(username, password, address):
    payload = getPayload("1", address)
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    for tryTims in range(10):

        r = requests.post(host, data=payload, cookies=ck)
        if r.status_code==200:
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            pageCount = rsp["pageCount"]
            print pageCount
            return pageCount
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host
            setSleepTimeInfo
            time.sleep(setSleepTime)
def getJsonWithTryTimes(host,payload,ck,method="post"):
    for tryTims in range(10):
        if method=="post":
            r = requests.post(host, data=payload, cookies=ck)
        else:
            r=requests.get(host, cookies=ck)
        if r.status_code==200:
            return r
        else:
            print u"服务器返回状态码错误。状态码为{statuscode}，返回内容为{rspcontent}".format(statuscode=r.status_code, rspcontent=r.text)
            print host.decode("utf-8")
            setSleepTimeInfo
            time.sleep(setSleepTime)
def getLockPwdCount(username, password, address):
    et.generateExcel()
    startTime = datetime.datetime.now()
    recordDay = startTime.strftime("%Y-%m-%d")
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    lockCountExcelRow = 1
    payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    # r = requests.post(host, data=payload, cookies=ck)
    r=getJsonWithTryTimes(host,payload,ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pagecount = rsp["pageCount"]
    rsp = rsp["resultList"]
    print "返回的页数为：".decode("utf-8"), pagecount
    # 为了调试方便，暂时设置为10
    # rowcount=10
    for page in range(1, pagecount + 1):
        print "开始遍历第%s页".decode("utf-8") % page

        lockInfoDic = {}
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]

            req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid

            r=getJsonWithTryTimes(req,"none",ck,"get")

            lock_rsp = r.text
            lock_rsp = json.loads(lock_rsp, encoding='utf-8')

            lock_rsp = lock_rsp["result"]

            if lock_rsp["lockList"]:
                # 先根据devid 获取门锁的预置密码，然后再根据devid找到id，再用id找到正常密码
                deviceID = lock_rsp["lockList"][0]["deviceId"]

                print "门锁设备ID:%s" % deviceID, "房间地址：%s" % i["hosueInfo"]["homeAddress"]
                nor_pwd_count = getPwdCountsInNormal(deviceID)
                # pre_pwd_count=getPwdCountsInPre(deviceID)
                # 计算密码已使用个数
                housekeepercount = nor_pwd_count["housekeeper"]
                renterCount = nor_pwd_count["renter"]
                # 从log Report中找到管家，租客，临时密码个数
                getDataSyncSql = "SELECT * FROM log_report WHERE device_id='" + deviceID + "' and report_time LIKE '%" + recordDay + "%' and url_path='status/data_sync'  ORDER BY report_time DESC"
                # 判断结果，如果没有，那么做一次密码同步

                getDataSyncResult = dbOperation(getDataSyncSql, db="danbay_task")

                if not getDataSyncResult:
                    for i in range(3):
                        syncPwdUrl = "http://www.danbay.cn/system/lock/socket/getPwdInfo?deviceId=%s" % deviceID
                        r = requests.get(syncPwdUrl, cookies=ck)
                        time.sleep(2)
                        getDataSyncResult = dbOperation(getDataSyncSql, db="danbay_task")
                        if getDataSyncResult:
                            break

                if getDataSyncResult:
                    syncContent = getDataSyncResult[0][5]
                    syncPayload = syncContent.split("payLoadString=")[1]
                    syncPayloadDict = json.loads(syncPayload)
                    gjCounts = syncPayloadDict["steward_cap"]
                    tmpCounts = syncPayloadDict["temp_cap"]
                    zuKeCounts = syncPayloadDict["user_cap"]
                    totalCounts = syncPayloadDict["total_cap"]

                    wb.get_sheet(0).write(lockCountExcelRow, 0, lock_rsp["roomName"])
                    wb.get_sheet(0).write(lockCountExcelRow, 1, gjCounts)
                    wb.get_sheet(0).write(lockCountExcelRow, 2, zuKeCounts)
                    wb.get_sheet(0).write(lockCountExcelRow, 3, tmpCounts)
                    wb.get_sheet(0).write(lockCountExcelRow, 4, renterCount)
                    wb.get_sheet(0).write(lockCountExcelRow, 5, housekeepercount)
                    wb.get_sheet(0).write(lockCountExcelRow, 6, gjCounts + zuKeCounts + tmpCounts + 1)
                    wb.get_sheet(0).write(lockCountExcelRow, 7, totalCounts)
                else:
                    wb.get_sheet(0).write(lockCountExcelRow, 0, lock_rsp["roomName"])
                    wb.get_sheet(0).write(lockCountExcelRow, 1, u'中控没上报相应日志，请手动检查')
                    wb.get_sheet(0).write(lockCountExcelRow, 2, u'中控没上报相应日志，请手动检查')
                    wb.get_sheet(0).write(lockCountExcelRow, 3, u'中控没上报相应日志，请手动检查')
                    wb.get_sheet(0).write(lockCountExcelRow, 4, renterCount)
                    wb.get_sheet(0).write(lockCountExcelRow, 5, housekeepercount)
                    wb.get_sheet(0).write(lockCountExcelRow, 6, u'中控没上报相应日志，无法统计，请手动检查')
                    wb.get_sheet(0).write(lockCountExcelRow, 7, u'中控没上报相应日志，无法统计，请手动检查')




                lockCountExcelRow = lockCountExcelRow + 1

        payload = {'pageNo': str(page + 1), 'pageSize': str(100), 'likeStr': address, "detailAddress": "", "floor": "",
                   "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'

        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        wb.save(u'test.xls')
    renameFile(address[0])
# def getLockPwdCountFromDB(username,password,address):
#     generateExcel()
#     lockCountExcelRow=1
#     payload = {'pageNo': '1', 'pageSize': '10', 'likeStr':address,"detailAddress":"","floor":"","spaceType":"","userId":"1"}
#     host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#     ck=getCookies(username,password)
#     r = requests.post(host, data=payload,cookies=ck)
#     rsp=json.loads(r.text)
#     rsp=rsp["result"]
#     pagecount=rsp["pageCount"]
#     rsp = rsp["resultList"]
#     print "返回的页数为：".decode("utf-8"),pagecount
#     # rowcount=10
#     for page in range(1,pagecount+1):
#         print "开始遍历第%s页".decode("utf-8")%page
#
#         lockInfoDic={}
#         for k in range(len(rsp)):
#             i=rsp[k]
#             houseid=i["hosueInfo"]["id"]
#             req="http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s"%houseid
#             r = requests.get(req,cookies=ck)
#             lock_rsp=r.text
#             lock_rsp=json.loads(lock_rsp,encoding='utf-8')
#             lock_rsp=lock_rsp["result"]
#             if lock_rsp["lockList"]:
#                 # 先根据devid 获取门锁的预置密码，然后再根据devid找到id，再用id找到正常密码
#                 # lockInfoDic[lock_rsp["lockList"][0]["deviceId"]]=lock_rsp["roomName"]
#                 deviceID=lock_rsp["lockList"][0]["deviceId"]
#
#                 print "门锁设备ID:%s"%deviceID,"房间地址：%s"%i["hosueInfo"]["homeAddress"]
#                 nor_pwd_count=getPwdCountsInNormal(deviceID)
#                 pre_pwd_count=getPwdCountsInPre(deviceID)
#                 # 计算密码已使用个数
#                 admincount=nor_pwd_count["admin"]
#                 housekeepercount=nor_pwd_count["housekeeper"]
#                 renterCount=nor_pwd_count["renter"]
#                 tmpCount=nor_pwd_count["tmp"]
#                 renterPreCount=pre_pwd_count["renter_pwd"]
#                 tmpPreCount=pre_pwd_count["tmp_pwd"]
#                 locktotalCount=int(admincount)+int(housekeepercount)+int(renterCount)+int(tmpCount)+int(renterPreCount)+int(tmpPreCount)
#                 writeExcel(lockCountExcelRow, 0, lock_rsp["roomName"])
#                 writeExcel(lockCountExcelRow, 1, nor_pwd_count["admin"])
#                 writeExcel(lockCountExcelRow, 2, nor_pwd_count["housekeeper"])
#                 writeExcel(lockCountExcelRow, 3, nor_pwd_count["renter"])
#                 writeExcel(lockCountExcelRow, 4, nor_pwd_count["tmp"])
#                 writeExcel(lockCountExcelRow, 5, pre_pwd_count["renter_pwd"])
#                 writeExcel(lockCountExcelRow, 6, pre_pwd_count["tmp_pwd"])
#                 writeExcel(lockCountExcelRow, 7, lock_rsp["lockList"][0]["deviceId"])
#                 writeExcel(lockCountExcelRow, 8, locktotalCount)
#                 lockCountExcelRow=lockCountExcelRow+1
#
#         payload = {'pageNo': str(page+1), 'pageSize': str(100), 'likeStr': address, "detailAddress": "", "floor": "",
#                    "spaceType": "",
#                    "userId": "1"}
#         host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]
#         rsp = rsp["resultList"]
#
#     renameFile(address[0])

def getAmmeterDeviceId(username, password, address):

    # 先根据地址判断，有没有采集器
    collectorSql="SELECT count(*) FROM collector_device cd,house_info hi WHERE cd.houseInfoId=hi.ID AND hi.homeAddress LIKE '%"+address+"%'"

    collectorCounts=dbOperation(collectorSql)
    if collectorCounts[0][0]==0:
        #无采集器
        et.generateMeterExcel()
        caiJiFlag=False
    else:
        #有采集器
        et.generateCaiJiQi()
        caiJiFlag = True

    excelRows = 1

    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'

    ck = getCookies(username, password)

    pageCount = getPageCount(username, password, address)

    for index in range(pageCount):
        pageNo = index + 1
        payload = getPayload(pageNo, address)
        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        # 判断是否有水电表

        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
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
                wb.get_sheet(0).write(excelRows, 0, homeaddress)

                # 设备类型，水保还是电表
                if meterType == "0":
                    if subType:
                        if subType == "1":
                            wb.get_sheet(0).write(excelRows, 1, u"热水")
                        elif subType == "0":
                            wb.get_sheet(0).write(excelRows, 1, u"冷水")
                    else:
                        wb.get_sheet(0).write(excelRows, 1, u"水表")
                elif meterType == "1":
                    wb.get_sheet(0).write(excelRows, 1, u"电表")
                if meterStatus == "1":
                    wb.get_sheet(0).write(excelRows, 2, u"离线", redStyle())
                elif meterStatus == "0":
                    wb.get_sheet(0).write(excelRows, 2, u"在线")
                #带采集器
                if caiJiFlag:
                    # 离线时长
                    collecterIDSql = "SELECT collectorId from energy_device where id=" + "\'" + str(
                        elecmeterID) + "\'" + ";"
                    collecterID = dbOperation(collecterIDSql)

                    collecterID = collecterID[0][0]
                    # 在采集器表中获取采集器的设备id
                    collecterSQl = "SELECT deviceId,macAddress FROM collector_device WHERE id=" + "\'" + str(collecterID) + "\'" + ";"
                    collecterSQlResult = dbOperation(collecterSQl)
                    collectDeviceID = collecterSQlResult[0][0] #采集器ID
                    collectDeviceAddress=collecterSQlResult[0][1]#采集器Mac

                    tt = checkDeviceOffLine(collectDeviceID)

                    wb.get_sheet(0).write(excelRows, 3, (tt.total_seconds() / (24 * 3600)))
                    wb.get_sheet(0).write(excelRows, 4, getDeviceOfflineCounts(collectDeviceID))

                    wb.get_sheet(0).write(excelRows, 5, getDeviceCounts(meterID))
                    shebeinmac = "SELECT address,deviceId,centerControlId from energy_device where id=" + "\'" + str(elecmeterID) + "\'" + ";"
                    shebeiMacResult = dbOperation(shebeinmac)
                    shebeiMac=shebeiMacResult[0][0]#设备表号
                    shebeiDevId=shebeiMacResult[0][1]#设备Id
                    centerCotrolID=shebeiMacResult[0][2]#中控ID

                    wb.get_sheet(0).write(excelRows, 6, collectDeviceAddress)
                    wb.get_sheet(0).write(excelRows, 7, collectDeviceID)
                    wb.get_sheet(0).write(excelRows, 8, shebeiMac)
                    wb.get_sheet(0).write(excelRows, 9, shebeiDevId)

                    # 中控deviceid
                    ControlAddress = "SELECT deviceId,macAddress,online FROM center_control WHERE id=" + "\'" + str(
                        centerCotrolID) + "\'" + ";"
                    centID = dbOperation(ControlAddress)  # 中控地址
                    wb.get_sheet(0).write(excelRows, 10, centID[0][1])
                    if centID[0][2] == 0:
                        wb.get_sheet(0).write(excelRows, 11, u'在线')
                    elif centID[0][2] == 1:
                        wb.get_sheet(0).write(excelRows, 11, u'离线', redStyle())
                    wb.get_sheet(0).write(excelRows, 12, centID[0][0])

                    print "完成房间:".decode("utf-8"), homeaddress, "的设备信息录入...".decode("utf-8")
                    excelRows = excelRows + 1

                if not caiJiFlag:
                    #不带采集器
                    # 计算离线总时长
                    # 根据deviceid找到对应的中控，并写入进去
                    # 如果带采集器，那么采集器的在线离线时间，就是设备的在线离线时间
                    getmeterControlIdSql = "SELECT centerControlId,address FROM energy_device WHERE deviceId=" + "\'" + meterID + "\'" + ";"
                    centerCotrolID = dbOperation(getmeterControlIdSql)[0][0]  # 中控DevId
                    meterMac = dbOperation(getmeterControlIdSql)[0][1]  # 中控Mac地址

                    tt = checkDeviceOffLine(meterID)

                    wb.get_sheet(0).write(excelRows, 3, (tt.total_seconds() / (24 * 3600)))
                    wb.get_sheet(0).write(excelRows, 4, getDeviceOfflineCounts(meterID))
                    wb.get_sheet(0).write(excelRows, 5, getDeviceCounts(meterID))
                    wb.get_sheet(0).write(excelRows, 6, meterMac)
                    wb.get_sheet(0).write(excelRows, 7, meterID)

                    # 中控deviceid
                    ControlAddress = "SELECT deviceId,macAddress,online FROM center_control WHERE id=" + "\'" + str(
                        centerCotrolID) + "\'" + ";"
                    centID = dbOperation(ControlAddress)  # 中控地址
                    wb.get_sheet(0).write(excelRows, 8, centID[0][1])
                    if centID[0][2] == 0:
                        wb.get_sheet(0).write(excelRows, 9, u'在线')
                    elif centID[0][2] == 1:
                        wb.get_sheet(0).write(excelRows, 9, u'离线', redStyle())
                    wb.get_sheet(0).write(excelRows, 10, centID[0][0])

                    print "完成房间:".decode("utf-8"), homeaddress, "的设备信息录入...".decode("utf-8")
                    excelRows = excelRows + 1

        wb.save(u'test.xls')



    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表状态_' + unicode(a, "utf-8") + u'.xls')
def getDeviceOfflineCounts(deviceID):
    checkStartTime=(datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
    checkEndTime=datetime.datetime.now().strftime("%Y-%m-%d")
    devOfflineCountsSql="SELECT count(*) FROM log_report WHERE device_id='"+deviceID+"' and url_path LIKE '%device/logout%' AND report_time BETWEEN '"+checkStartTime+" 00:00:00' AND  '"+checkEndTime+" 00:00:00'  ORDER BY report_time DESC "
    devOfflineCountsResult=dbOperation(devOfflineCountsSql,db="danbay_task")
    devOfflineCountsLen=devOfflineCountsResult[0][0]
    return devOfflineCountsLen
def getDeviceCounts(devId):
    meterCountsSql = "SELECT meterCount FROM energy_device WHERE deviceid='" + devId + "'"
    meterCountResult = dbOperation(meterCountsSql)
    return meterCountResult[0][0]

def getShuiDianInfoFromDB(address):
    excelRows = 1
    et.generateShuiDianExcelFromDB()
    # address="交叉口"
    queryResultSql = "SELECT ID,homeAddress FROM house_info WHERE homeAddress LIKE '%" + address + "%' and deleteState !=1;"
    homeAddresssAndIDResult = dbOperation(queryResultSql)
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
    for addrAndId in range(0, len(homeAddresssAndIDResult)):
        deviceQuerySql = "SELECT type,online,address,deviceId,meterCount,offlineTime,onlineTime,centerControlId,collectorId FROM energy_device where houseInfoId = " + str(
            homeAddresssAndIDResult[addrAndId][0]) + ";"
        queryDeviceResult = dbOperation(deviceQuerySql)
        # for queryDeviceResultItem in range(len(queryDeviceResult)):
        #     if queryDeviceResult[queryDeviceResultItem]:
        #         writeExcel(excelRows, 0, homeAddresssAndIDResult[addrAndId][1])
        #         # print queryDeviceResult[queryDeviceResultItem]
        #         if queryDeviceResult[queryDeviceResultItem][0]=='0':  #设备类型，水表
        #             writeExcel(excelRows,1, u'水表')
        #         elif queryDeviceResult[queryDeviceResultItem][0]=='1':#设备类型，电表表
        #             writeExcel(excelRows, 1, u'电表')
        #         if queryDeviceResult[queryDeviceResultItem][1]==0: #设备状态，在线 ，离线
        #             writeExcel(excelRows, 2, u'在线')
        #         else:
        #             writeExcel(excelRows, 2, u'离线')
        #         writeExcel(excelRows, 3, queryDeviceResult[queryDeviceResultItem][2]) #mac addres
        #         writeExcel(excelRows, 4, queryDeviceResult[queryDeviceResultItem][3])#deviceid
        #         writeExcel(excelRows, 5, queryDeviceResult[queryDeviceResultItem][4])#表头读数
        #         writeExcel(excelRows, 6, queryDeviceResult[queryDeviceResultItem][5],3)#离线时间
        #         writeExcel(excelRows, 7, queryDeviceResult[queryDeviceResultItem][6],3)#在线时间
        #         if queryDeviceResult[0][8]:
        #             writeExcel(excelRows, 9, queryDeviceResult[queryDeviceResultItem][8])  # 采集器id
        #         # writeExcel(excelRows, 8, queryDeviceResult[queryDeviceResultItem][7]) #中控id
        #         # 获取中控deviceId， 中控Mac，中控状态，中控关联地址
        #         # print queryDeviceResult[queryDeviceResultItem][7]
        #         getShuiDianCCSql="SELECT online,deviceId,macAddress,version,address from center_control WHERE id="+str(queryDeviceResult[queryDeviceResultItem][7])+";"
        #         getShuiDianCCSqlResult=dbOperation(getShuiDianCCSql)
        #
        #         if getShuiDianCCSqlResult[0][0]==0: #中控在线 离线
        #             writeExcel(excelRows, 8, u'在线')
        #         else:
        #             writeExcel(excelRows, 8, u'离线')
        #         writeExcel(excelRows, 9, getShuiDianCCSqlResult[0][1]) # 中控id
        #         writeExcel(excelRows, 10, getShuiDianCCSqlResult[0][2])#中控mac
        #         writeExcel(excelRows, 11, getShuiDianCCSqlResult[0][3])#中控关联地址
        #         excelRows=excelRows+1
        #         print homeAddresssAndIDResult[addrAndId][1]
        # 如果查询结果为空，则表示没有该设备，不用处理

        #######################################################更新写操作##############################
        for queryDeviceResultItem in range(len(queryDeviceResult)):
            if queryDeviceResult[queryDeviceResultItem]:
                wb.get_sheet(0).write(excelRows, 0, homeAddresssAndIDResult[addrAndId][1])
                if queryDeviceResult[queryDeviceResultItem][0] == '0':  # 设备类型，水表
                    wb.get_sheet(0).write(excelRows, 1, u'水表')
                elif queryDeviceResult[queryDeviceResultItem][0] == '1':  # 设备类型，电表表
                    wb.get_sheet(0).write(excelRows, 1, u'电表')
                if queryDeviceResult[queryDeviceResultItem][1] == 0:  # 设备状态，在线 ，离线
                    wb.get_sheet(0).write(excelRows, 2, u'在线')
                else:
                    wb.get_sheet(0).write(excelRows, 2, u'离线', redStyle())
                wb.get_sheet(0).write(excelRows, 3, queryDeviceResult[queryDeviceResultItem][2])
                wb.get_sheet(0).write(excelRows, 4, queryDeviceResult[queryDeviceResultItem][3])
                wb.get_sheet(0).write(excelRows, 5, queryDeviceResult[queryDeviceResultItem][4])
                wb.get_sheet(0).write(excelRows, 6, queryDeviceResult[queryDeviceResultItem][5], date_format)
                wb.get_sheet(0).write(excelRows, 7, queryDeviceResult[queryDeviceResultItem][6], date_format)
                # if queryDeviceResult[0][8]:
                #     writeExcel(excelRows, 9, queryDeviceResult[queryDeviceResultItem][8])  # 采集器id

                # 获取中控deviceId， 中控Mac，中控状态，中控关联地址
                # print queryDeviceResult[queryDeviceResultItem][7]
                getShuiDianCCSql = "SELECT online,deviceId,macAddress,version,address from center_control WHERE id=" + str(
                    queryDeviceResult[queryDeviceResultItem][7]) + ";"
                getShuiDianCCSqlResult = dbOperation(getShuiDianCCSql)

                if getShuiDianCCSqlResult[0][0] == 0:  # 中控在线 离线

                    wb.get_sheet(0).write(excelRows, 8, u'在线')
                else:
                    wb.get_sheet(0).write(excelRows, 8, u'离线', redStyle())
                wb.get_sheet(0).write(excelRows, 9, getShuiDianCCSqlResult[0][1])
                wb.get_sheet(0).write(excelRows, 10, getShuiDianCCSqlResult[0][2])
                wb.get_sheet(0).write(excelRows, 11, getShuiDianCCSqlResult[0][3])
                excelRows = excelRows + 1
                print homeAddresssAndIDResult[addrAndId][1]
    wb.save(u'test.xls')

    # 根据房源地址，找到用户信息，然后再找到合同号，根据合同号，找出所有的设备信息，然后写入Excel中

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表状态_' + unicode(a, "utf-8") + u'.xls')

def LockOnlineStatus(username, password, address):
    excelRows = 1
    et.generateLockInfoExcel()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
    payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    r=getJsonWithTryTimes(host,payload,ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pageCount = rsp["pageCount"]
    for index in range(pageCount):
        pageNo = index + 1
        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        payload = getPayload(pageNo, address)
        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            locklist = devicersp["lockList"]
            gatewalist = devicersp["gatewayList"]
            lockRoomName = devicersp["roomName"]

            for deviceIndex in range(len(locklist)):
                lock = locklist[deviceIndex]
                lcokDeviceId = lock["deviceId"]
                lcokOnlineStatus = lock["onlineStatus"]
                wb.get_sheet(0).write(excelRows, 0, lockRoomName)
                if lcokOnlineStatus == "0":
                    wb.get_sheet(0).write(excelRows, 1, u"在线")
                else:
                    wb.get_sheet(0).write(excelRows, 1, u"离线", redStyle())

                wb.get_sheet(0).write(excelRows, 2, lcokDeviceId)
                # 门锁mac
                lockMacAddressSql = "SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                lockMacAddressResult = dbOperation(lockMacAddressSql)
                lockMacAddress = lockMacAddressResult[0][0]
                wb.get_sheet(0).write(excelRows, 3, lockMacAddress)

                # 从数据库中获取中控ID-----
                lockCenterControlIDSql = "SELECT centerControl from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                lockCenterControlID = dbOperation(lockCenterControlIDSql)
                # 循环中控，找到该门锁的id
                writeExcelFlag = False

                if len(gatewalist) == 1:
                    writeExcelFlag = True
                    gwonlineStatus = gatewalist[0]["onlineStatus"]
                    gwid = gatewalist[0]["id"]
                else:
                    for gateway in gatewalist:
                        gwid = gateway["id"]
                        # 如果门锁的中控ID跟当前获取到的中控id一致，那么记录该中控在下状态
                        if gwid == lockCenterControlID[0][0]:
                            writeExcelFlag = True
                            gwonlineStatus = gateway["onlineStatus"]
                    # 中控mac
                    # 中控在线状态

                if writeExcelFlag:

                    if gwonlineStatus == "1":
                        # 离线
                        wb.get_sheet(0).write(excelRows, 4, u'离线', redStyle())
                    elif gwonlineStatus == "0":

                        wb.get_sheet(0).write(excelRows, 4, u'在线')

                    centerControlMacSql = "SELECT deviceId,macAddress,version from center_control where id=" + "\'" + str(
                        gwid) + "\'" + ";"
                    centerControlMacSqlResult = dbOperation(centerControlMacSql)
                    centerControlMac = centerControlMacSqlResult[0][1]
                    centerControlDevID = centerControlMacSqlResult[0][0]
                    centerControlVersion = centerControlMacSqlResult[0][2]

                    wb.get_sheet(0).write(excelRows, 5, centerControlDevID)
                    wb.get_sheet(0).write(excelRows, 6, centerControlMac)
                    wb.get_sheet(0).write(excelRows, 7, centerControlVersion)

                excelRows = excelRows + 1
                print "完成房间:".decode("utf-8"), lockRoomName, "的设备信息录入...".decode("utf-8")
        wb.save(u'test.xls')

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address[0], "utf-8") + u'_门锁在线离线_' + unicode(a, "utf-8") + u'.xls')


# def DelLockPwd(username, password, address):
#     excelRows = 1
#     generateLockInfoExcel()
#     rb = xlrd.open_workbook("test.xls", formatting_info=True)
#     wb = copy(rb)
#
#     payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
#                "userId": "1"}
#     host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#     ck = getCookies(username, password)
#     r=getJsonWithTryTimes(host,payload,ck)
#     rsp = json.loads(r.text)
#     rsp = rsp["result"]
#     pageCount = rsp["pageCount"]
#     for index in range(pageCount):
#         pageNo = index + 1
#         print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
#         payload = getPayload(pageNo, address)
#         r=getJsonWithTryTimes(host,payload,ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]
#         rsp = rsp["resultList"]
#         for k in range(len(rsp)):
#             i = rsp[k]
#             houseid = i["hosueInfo"]["id"]
#             devicersp = getJson(ck, houseid)
#             devicersp = json.loads(devicersp, encoding='utf-8')
#             devicersp = devicersp["result"]
#             locklist = devicersp["lockList"]
#             gatewalist = devicersp["gatewayList"]
#             lockRoomName = devicersp["roomName"]
#
#             for deviceIndex in range(len(locklist)):
#                 lock = locklist[deviceIndex]
#                 lcokDeviceId = lock["deviceId"]
#                 lcokOnlineStatus = lock["onlineStatus"]
#                 # 循环获取门锁租客密码数量，然后根据开门记录，删除密码
#                 wb.get_sheet(0).write(excelRows, 0, lockRoomName)
#
#                 # 获取门锁密码数量
#                 print lcokDeviceId
#                 if lcokOnlineStatus == "0":
#                     wb.get_sheet(0).write(excelRows, 1, u"在线")
#                     # 如果在线，则删除密码
#
#                 else:
#                     # 如果离线，则写设备离线，无法删除密码
#                     wb.get_sheet(0).write(excelRows, 1, u"离线", redStyle())
#
#                 wb.get_sheet(0).write(excelRows, 2, lcokDeviceId)
#                 # # 门锁mac
#                 # lockMacAddressSql="SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
#                 # lockMacAddressResult=dbOperation(lockMacAddressSql)
#                 # lockMacAddress=lockMacAddressResult[0][0]
#                 # # writeExcel(excelRows, 3, lockMacAddress)
#                 # wb.get_sheet(0).write(excelRows, 3, lockMacAddress)
#
#                 # # 从数据库中获取中控ID-----
#                 # lockCenterControlIDSql = "SELECT centerControl from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
#                 # lockCenterControlID = dbOperation(lockCenterControlIDSql)
#                 # # 循环中控，找到该门锁的id
#                 # writeExcelFlag=False
#                 #
#                 # if len(gatewalist)==1:
#                 #     writeExcelFlag = True
#                 #     gwonlineStatus=gatewalist[0]["onlineStatus"]
#                 #     gwid=gatewalist[0]["id"]
#                 # else:
#                 #     for gateway in gatewalist:
#                 #         gwid = gateway["id"]
#                 #         # 如果门锁的中控ID跟当前获取到的中控id一致，那么记录该中控在下状态
#                 #         if gwid == lockCenterControlID[0][0]:
#                 #             writeExcelFlag = True
#                 #             gwonlineStatus = gateway["onlineStatus"]
#                 #     # 中控mac
#                 #     # 中控在线状态
#                 #
#                 # if writeExcelFlag:
#                 #
#                 #     if gwonlineStatus == "1":
#                 #     # 离线
#                 #     #     writeExcel(excelRows, 4, u'离线',2)
#                 #         wb.get_sheet(0).write(excelRows, 4, u'离线',redStyle())
#                 #     elif gwonlineStatus == "0":
#                 #         # writeExcel(excelRows, 4, u'在线')
#                 #
#                 #         wb.get_sheet(0).write(excelRows, 4, u'在线')
#                 #
#                 #     centerControlMacSql="SELECT deviceId,macAddress from center_control where id=" + "\'" + str(gwid) + "\'" + ";"
#                 #     centerControlMacSqlResult=dbOperation(centerControlMacSql)
#                 #     centerControlMac=centerControlMacSqlResult[0][1]
#                 #     centerControlDevID=centerControlMacSqlResult[0][0]
#                 #     # print centerControlMac,centerControlDevID,gwid
#                 #
#                 #     # writeExcel(excelRows, 5, centerControlDevID)
#                 #     wb.get_sheet(0).write(excelRows, 5, centerControlDevID)
#                 #     # writeExcel(excelRows, 6, centerControlMac)
#                 #     wb.get_sheet(0).write(excelRows, 6, centerControlMac)
#
#                 excelRows = excelRows + 1
#         wb.save(u'test.xls')

    # print  "done"
    # a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    #
    # os.rename("test.xls", unicode(address[0], "utf-8") + u'_门锁在线离线_' + unicode(a, "utf-8") + u'.xls')

def getFenSanHomeAddr():

    getalladdresssql = "SELECT location FROM homesourcelist WHERE homeSourceProviderId=163"
    getalladdress = dbOperation(getalladdresssql, db='danbay_projects')
    return  getalladdress

def fensanLockOnlineStatus(username, password):
    # 获取所有房源地址
    excelRows = 1
    et.generateLockInfoExcel()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    getalladdress=getFenSanHomeAddr()

    for homeaddr in getalladdress:
        address = homeaddr[0]
        payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "",
                   "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
        ck = getCookies(username, password)
        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        for index in range(pageCount):
            pageNo = index + 1
            print u"正在查询第%s页" % pageNo
            payload = getPayload(pageNo, address)
            r=getJsonWithTryTimes(host,payload,ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            rsp = rsp["resultList"]
            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                devicersp = getJson(ck, houseid)
                devicersp = json.loads(devicersp, encoding='utf-8')
                devicersp = devicersp["result"]
                locklist = devicersp["lockList"]

                gatewalist = devicersp["gatewayList"]
                lockRoomName = devicersp["roomName"]
                for deviceIndex in range(len(locklist)):
                    lock = locklist[deviceIndex]
                    lcokDeviceId = lock["deviceId"]
                    lcokOnlineStatus = lock["onlineStatus"]
                    wb.get_sheet(0).write(excelRows, 0, lockRoomName)
                    if lcokOnlineStatus == "0":
                        wb.get_sheet(0).write(excelRows, 1, u"在线")
                    else:
                        wb.get_sheet(0).write(excelRows, 1, u"离线", redStyle())
                    wb.get_sheet(0).write(excelRows, 2, lcokDeviceId)
                    # 门锁mac
                    lockMacAddressSql = "SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                    lockMacAddressResult = dbOperation(lockMacAddressSql)
                    lockMacAddress = lockMacAddressResult[0][0]
                    wb.get_sheet(0).write(excelRows, 3, lockMacAddress)

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
                                wb.get_sheet(0).write(excelRows, 4, u'离线', redStyle())
                            elif gwonlineStatus == "0":
                                wb.get_sheet(0).write(excelRows, 4, u'在线')
                            centerControlMacSql = "SELECT deviceId,macAddress,version from center_control where id=" + "\'" + str(
                                gwid) + "\'" + ";"
                            centerControlMacSqlResult = dbOperation(centerControlMacSql)
                            centerControlMac = centerControlMacSqlResult[0][1]
                            centerControlDevID = centerControlMacSqlResult[0][0]
                            centerControlVersion = centerControlMacSqlResult[0][2]

                            wb.get_sheet(0).write(excelRows, 5, centerControlDevID)
                            wb.get_sheet(0).write(excelRows, 6, centerControlMac)
                            wb.get_sheet(0).write(excelRows, 7, centerControlVersion)
                        # 在线
                    # 写中控id
                    print lockRoomName
                    excelRows = excelRows + 1

        wb.save(u'test.xls')
    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))

    os.rename("test.xls", u'分散式公寓' + u'_门锁在线离线_' + unicode(a, "utf-8") + u'.xls')
    # 碧桂园门锁预置密码

def getFenSanCenterControlInfo(username, password):
    et.generateCenterControlInfoAtDeviceCenter()
    excelRowCount = 1
    addrs=getFenSanHomeAddr()

    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
    ck = getCookies(username, password)
    for singleAddr in addrs:
        address=singleAddr[0]
        pageCount = centerControlGetPageCount(username, password, address)
        # 获取中控id
        for index in range(pageCount):
            pageNo = index + 1
            payload = getCenterControlPayload(pageNo, address)
            r = getJsonWithTryTimes(host, payload, ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]["result"]
            rsp = rsp["resultList"]
            for k in range(len(rsp)):
                resultrow = rsp[k]
                ccid = resultrow["id"]
                ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
                ccMac = dbOperation(ccMacSql)
                ccMac = ccMac[0][0]
                print ccMac,resultrow["houseAddress"]
                houseAddress = resultrow["houseAddress"]
                centerControlstatus = resultrow["status"]
                centerCtrolVersion = resultrow["version"]

                wb.get_sheet(0).write(excelRowCount, 0, ccMac)
                wb.get_sheet(0).write(excelRowCount, 1, centerCtrolVersion)

                if centerControlstatus == "1":
                    wb.get_sheet(0).write(excelRowCount, 2, u"离线", redStyle())
                else:
                    wb.get_sheet(0).write(excelRowCount, 2, u"在线")

                writeExcel(excelRowCount, 3, houseAddress)
                wb.get_sheet(0).write(excelRowCount, 3, houseAddress)
                excelRowCount = excelRowCount + 1

    wb.save(u'test.xls')

    print  "done"

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", u'碧桂园分散式公寓_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')

def BGYLockPwd(username, password):
    # 获取所有房源地址
    excelRows = 1
    et.generateLockPwdInfoExcel()
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
        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        print "查询地址：" + address
        print "返回总页数：" + str(pageCount)
        for index in range(pageCount):
            pageNo = index + 1
            print "正在查询第%s页" % pageNo
            payload = getPayload(pageNo, address)
            r=getJsonWithTryTimes(host,payload,ck)
            rsp = json.loads(r.text)
            rsp = rsp["result"]
            rsp = rsp["resultList"]
            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                devicersp = getJson(ck, houseid)
                devicersp = json.loads(devicersp, encoding='utf-8')
                devicersp = devicersp["result"]
                locklist = devicersp["lockList"]
                gatewalist = devicersp["gatewayList"]
                lockRoomName = devicersp["roomName"]
                for deviceIndex in range(len(locklist)):
                    lock = locklist[deviceIndex]
                    lcokDeviceId = lock["deviceId"]
                    lcokOnlineStatus = lock["onlineStatus"]
                    writeExcel(excelRows, 0, lockRoomName)
                    writeExcel(excelRows, 1, lcokDeviceId)
                    # 获取预置租客密码
                    # getZuKePwd="SELECT * from lock_pre_password WHERE dev_id='523fd1b9a21226d418e0776d14a8abee' AND psw_type=3"
                    getZuKePwdSql = "SELECT psw_value from lock_pre_password WHERE dev_id=" + "\'" + lcokDeviceId + "\'" + "AND psw_type=3 AND delete_state=0" + ";"
                    zuKePwd = dbOperation(getZuKePwdSql)
                    try:
                        zkPwd = zuKePwd[0][0]
                        writeExcel(excelRows, 2, zkPwd)
                    except:
                        writeExcel(excelRows, 2, u'没有预置租客密码')

                    # 获取预置临时密码1
                    getPreTempPwdSql = "SELECT psw_value from lock_pre_password WHERE dev_id=" + "\'" + lcokDeviceId + "\'" + "AND psw_type=0 AND delete_state=0" + ";"
                    PreTempPwd = dbOperation(getPreTempPwdSql)
                    # print  PreTempPwd
                    for ptpIndex in range(len(PreTempPwd)):
                        writeExcel(excelRows, 3 + ptpIndex, PreTempPwd[ptpIndex][0])

                    print lockRoomName
                    excelRows = excelRows + 1

    print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", u'碧桂圆分散式公寓' + u'_门锁预置密码_' + unicode(a, "utf-8") + u'.xls')


# def getCenterControlInfo(username, password,address):
#     generateCenterControlinfo()
#     addedCenterID = []
#     host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#     ck = getCookies(username, password)
#     pageCount = getPageCount(username, password, address)
#     rb = xlrd.open_workbook("test.xls", formatting_info=True)
#     wb = copy(rb)
#     date_format = xlwt.XFStyle()
#     date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
#     # 获取中控id
#     for index in range(pageCount):
#         pageNo = index + 1
#         payload = getPayload(pageNo, address)
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]
#         rsp = rsp["resultList"]
#         for k in range(len(rsp)):
#             i = rsp[k]
#             houseid = i["hosueInfo"]["id"]
#             # # 通过 houseid 获取某个房间的设备
#             req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
#             r = requests.get(req, cookies=ck)
#             # devicersp = r.text
#             devicersp=getJson(ck, houseid)
#             # devicersp = getJson(ck, houseid)
#
#             devicersp = json.loads(devicersp, encoding='utf-8')
#             devicersp = devicersp["result"]
#             gatewayList = devicersp["gatewayList"]
#
#
#             for deviceIndex in range(len(gatewayList)):
#                 gw = gatewayList[deviceIndex]
#                 homeaddress = getHomeAddress(houseid)
#                 gwID = gw["id"]
#                 if gwID not in addedCenterID:
#                     addedCenterID.append(gwID)
#             print addedCenterID
#     ck = getCookies(username, password)
#     print "中控个数为：",len(addedCenterID)
#     for controlIDIndex in range(len(addedCenterID)):
#         devicersp = getCenterControlJson(ck,addedCenterID[controlIDIndex])
#         devicersp = json.loads(devicersp, encoding='utf-8')
#         devicersp = devicersp["result"]
#         devicersp=devicersp["basicInfo"]
#         centerContronDevmacAddress = devicersp["macAddress"]
#         centerContronDevVersion = devicersp["deviceModel"]
#         centerContronDevID=devicersp["deviceId"]
#         centerContronDevaddress=devicersp["address"]
#
#         writeExcel(controlIDIndex+1,0,centerContronDevmacAddress)
#         writeExcel(controlIDIndex+1,1,centerContronDevVersion)
#         writeExcel(controlIDIndex+1,2,centerContronDevID)
#         writeExcel(controlIDIndex+1,3,centerContronDevaddress)
#
#     print  "done"
#     a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
#     os.rename("test.xls", unicode(address, "utf-8") + u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')

# 通过设备中心获取中控信息

def getCenterControlInfoByDeviceCenter(username, password, address):
    et.generateCenterControlInfoAtDeviceCenter()

    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'
    host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'

    ck = getCookies(username, password)
    pageCount = centerControlGetPageCount(username, password, address)
    # 获取中控id

    excelRowCount = 1
    for index in range(pageCount):

        pageNo = index + 1

        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        payload = getCenterControlPayload(pageNo, address)
        r=getJsonWithTryTimes(host,payload,ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]["result"]
        rsp = rsp["resultList"]

        for k in range(len(rsp)):
            resultrow = rsp[k]
            ccid = resultrow["id"]
            ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
            ccMac = dbOperation(ccMacSql)
            ccMac = ccMac[0][0]
            print ccMac
            houseAddress = resultrow["houseAddress"]
            centerControlstatus = resultrow["status"]
            centerCtrolVersion = resultrow["version"]

            wb.get_sheet(0).write(excelRowCount, 0, ccMac)
            wb.get_sheet(0).write(excelRowCount, 1, centerCtrolVersion)

            if centerControlstatus == "1":
                wb.get_sheet(0).write(excelRowCount, 2, u"离线", redStyle())
            else:
                wb.get_sheet(0).write(excelRowCount, 2, u"在线")

            # writeExcel(excelRowCount, 3, houseAddress)
            wb.get_sheet(0).write(excelRowCount, 3, houseAddress)
            excelRowCount = excelRowCount + 1

        wb.save(u'test.xls')

    print "done"

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address[0], "utf-8") + u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')


def checkLockDataSync(username, password, addr, checkOffLien, singleFlag=1):
    excelRows = 1
    et.generateCheckLockSync()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    homesourceListSql = "SELECT location from homesourcelist WHERE deleteState =0"
    homesourceList = dbOperation(homesourceListSql, db="danbay_projects")

    for homeAddr in homesourceList:
        address = homeAddr[0]
        if singleFlag == 1:
            address = addr
            stopFullRun = 1
        else:
            stopFullRun = 5
        payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "",
                   "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
        ck = getCookies(username, password)
        getHouseInfoByConditionjsonResult = getCommonJson(host, payload, ck, "post")
        rsp = json.loads(getHouseInfoByConditionjsonResult)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        for index in range(pageCount):
            pageNo = index + 1
            print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
            payload = getPayload(pageNo, address)
            getSpecialPageJsonResult = getCommonJson(host, payload, ck, "post")
            rsp = json.loads(getSpecialPageJsonResult)
            rsp = rsp["result"]
            rsp = rsp["resultList"]

            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                devicersp = getJson(ck, houseid)
                devrsp = json.loads(devicersp, encoding='utf-8')

                devicersp = devrsp["result"]
                try:
                    locklist = devicersp["lockList"]
                except:
                    print devicersp
                    locklist = 0
                if locklist:
                    lockRoomName = devicersp["roomName"]

                    # log.Log("{optTime}--获取到--{room}--的状态码为--{houseData}".format(optTime=datetime.datetime.now(),room=lockRoomName,houseData=devrsp["status"] + devrsp["message"]))
                    for deviceIndex in range(len(locklist)):
                        lock = locklist[deviceIndex]
                        lcokDeviceId = lock["deviceId"]

                        if checkOffLien == 1:
                            tt = checkDeviceOffLine(lcokDeviceId)
                            print lockRoomName, tt
                        else:
                            lockMacAddressSql = "SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                            lockMacAddressResult = dbOperation(lockMacAddressSql)
                            lockMacAddress = lockMacAddressResult[0][0]
                            checkStartTime = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime(
                                "%Y-%m-%d")

                            checkSql = "SELECT count(*)  FROM log_report WHERE device_id=\'" + lcokDeviceId + "\'  AND url_path='status/data_sync' AND report_time like \'%" + checkStartTime + "%' ORDER BY res_time DESC"
                            checkResult = dbOperation(checkSql, db="danbay_task")
                            # log.Log("查询数据库完毕！")
                            print(u"完成对设备的检查%s"%lcokDeviceId)
                            # prin/t checkTime + "同步次数为:" + str(checkResult[0][0])+lockRoomName
                            if checkResult[0][0] > 100:
                                wb.get_sheet(0).write(excelRows, 0, lockRoomName)
                                wb.get_sheet(0).write(excelRows, 1, lockMacAddress)
                                wb.get_sheet(0).write(excelRows, 2, lcokDeviceId)
                                wb.get_sheet(0).write(excelRows, 3,
                                                      (checkStartTime + "同步次数为:" + str(checkResult[0][0]).decode("utf-8")))
                                print (checkStartTime+ "同步次数为:" + str(checkResult[0][0])).decode("utf-8") , lockRoomName.decode("utf-8")
                                excelRows = excelRows + 1

            wb.save(u'test.xls')

        if stopFullRun == 1:
            break
            # print  "done"
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    if singleFlag == 1:
        os.rename("test.xls", address.decode("utf-8") + u"门锁数据同步次数统计_" + unicode(a, "utf-8") + u'.xls')
    else:
        os.rename("test.xls", u"全网门锁数据同步次数统计_" + unicode(a, "utf-8") + u'.xls')


def checkDeviceOffline(username, password, addr, checkOffLien, singleFlag=1):
    excelRows = 1
    generateDeviceOffline()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    homesourceListSql = "SELECT location from homesourcelist WHERE deleteState =0"
    homesourceList = dbOperation(homesourceListSql, db="danbay_projects")

    for homeAddr in homesourceList:
        address = homeAddr[0]
        if singleFlag == 1:
            address = addr
            stopFullRun = 1
        else:
            stopFullRun = 5
        payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "",
                   "spaceType": "",
                   "userId": "1"}
        host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
        ck = getCookies(username, password)
        getHouseInfoByConditionjsonResult = getCommonJson(host, payload, ck, "post")
        rsp = json.loads(getHouseInfoByConditionjsonResult)
        rsp = rsp["result"]
        pageCount = rsp["pageCount"]
        for index in range(pageCount):
            pageNo = index + 1
            print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
            payload = getPayload(pageNo, address)
            getSpecialPageJsonResult = getCommonJson(host, payload, ck, "post")
            rsp = json.loads(getSpecialPageJsonResult)
            rsp = rsp["result"]
            rsp = rsp["resultList"]

            for k in range(len(rsp)):
                i = rsp[k]
                houseid = i["hosueInfo"]["id"]
                devicersp = getJson(ck, houseid)
                devrsp = json.loads(devicersp, encoding='utf-8')

                devicersp = devrsp["result"]

                try:
                    locklist = devicersp["lockList"]
                except:
                    print devicersp
                    locklist = 0
                if locklist:
                    lockRoomName = devicersp["roomName"]

                    # log.Log("{optTime}--获取到--{room}--的状态码为--{houseData}".format(optTime=datetime.datetime.now(),room=lockRoomName,houseData=devrsp["status"] + devrsp["message"]))
                    for deviceIndex in range(len(locklist)):
                        lock = locklist[deviceIndex]
                        lcokDeviceId = lock["deviceId"]
                        lockMacAddressSql = "SELECT macAddress from device_info WHERE deviceId=" + "\'" + lcokDeviceId + "\'" + ";"
                        lockMacAddressResult = dbOperation(lockMacAddressSql)
                        lockMacAddress = lockMacAddressResult[0][0]

                        if checkOffLien == 1:
                            tt = checkDeviceOffLine(lcokDeviceId)
                            wb.get_sheet(0).write(excelRows, 0, lockRoomName)
                            wb.get_sheet(0).write(excelRows, 1, lockMacAddress)
                            wb.get_sheet(0).write(excelRows, 2, lcokDeviceId)
                            wb.get_sheet(0).write(excelRows, 3, (tt.total_seconds() / (24 * 3600)))
                            excelRows = excelRows + 1

                    try:
                        meterList = devicersp["meterList"]
                    except:
                        print devicersp
                        meterList = 0
                    if meterList:
                        pass

                    wb.save(u'test.xls')


def checkAllLockSyncsInDB():
    excelRows = 1
    loopTime = 1
    generateCheckLockSync()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    # "根据  用户id 获取 合同 id
    # 根据合同id 获取合同号，如果合同号一样，则不添加
    #     获取该合同号下，是否有门锁
    #         如果有门锁，遍历门锁的同步次数"

    getContractinfoSql = "SELECT DISTINCT contract_info.contractNum FROM mc_user,contract_info WHERE contract_info.`user`=mc_user.ID"
    getContractInfo = dbOperation(getContractinfoSql)
    # 遍历合同号
    for contract in getContractInfo:
        print contract[0]
        getLockSql = "SELECT * FROM device_info WHERE contractNum='" + contract[0] + "'"
        getLockResult = dbOperation(getLockSql)
        if getLockResult:
            for singleLock in getLockResult:
                # print singleLock[11],singleLock[12],singleLock[35]
                if singleLock[35]:
                    homeAddrSql = "SELECT homeAddress FROM house_info WHERE ID=" + str(
                        singleLock[35])  # singleLock[35]=houseINfo ID
                    homeAddr = dbOperation(homeAddrSql)
                    checkTime = "2018-05-04"
                    checkSql = "SELECT count(*)  FROM log_report WHERE device_id=\'" + singleLock[
                        11] + "\'  AND url_path='status/data_sync' AND report_time like \'%" + checkTime + "%' ORDER BY res_time DESC"
                    checkResult = dbOperation(checkSql, db="danbay_task")
                    loopTime = loopTime + 1
                    print "数据库查询完毕", loopTime
                    if checkResult[0][0] > 100:
                        wb.get_sheet(0).write(excelRows, 0, homeAddr[0][0])  # 房间地址
                        wb.get_sheet(0).write(excelRows, 1, singleLock[12])  # mac地址
                        wb.get_sheet(0).write(excelRows, 2, singleLock[11])  # deviceID
                        wb.get_sheet(0).write(excelRows, 3,
                                              (checkTime + "同步次数为:   " + str(checkResult[0][0]).decode("utf-8")))
                        print checkTime + "同步次数为:" + str(checkResult[0][0]), homeAddr[0][0]
                        excelRows = excelRows + 1
                        print
        wb.save(u'test.xls')

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", u"全网门锁数据同步次数统计_" + unicode(a, "utf-8") + u'.xls')


def checkLockSyncWithThreadsFuntion(treadId, checkTime, contractInfoList):
    # "根据  用户id 获取 合同 id
    # 根据合同id 获取合同号，如果合同号一样，则不添加
    #     获取该合同号下，是否有门锁
    #         如果有门锁，遍历门锁的同步次数"
    # 先多线程找到结果，存放到list中，然后再从list中写古Excel

    # 先将记录存放到列表中，全部统计完之后，再返回出去，然后写入Excel
    searchResultList = []
    loopTime = 1

    # 遍历合同号
    for contract in contractInfoList:
        print contract[0]
        getLockSql = "SELECT * FROM device_info WHERE contractNum='" + contract[0] + "'"
        getLockResult = dbOperation(getLockSql)
        if getLockResult:
            for singleLock in getLockResult:
                singleLockRecordList = []
                if singleLock[35]:
                    homeAddrSql = "SELECT homeAddress FROM house_info WHERE ID=" + str(
                        singleLock[35])  # singleLock[35]=houseINfo ID
                    homeAddr = dbOperation(homeAddrSql)
                    checkSql = "SELECT count(*)  FROM log_report WHERE device_id=\'" + singleLock[
                        11] + "\'  AND url_path='status/data_sync' AND report_time like \'%" + checkTime + "%' ORDER BY res_time DESC"
                    checkResult = dbOperation(checkSql, db="danbay_task")
                    loopTime = loopTime + 1
                    print str(treadId) + "数据库查询完毕", loopTime
                    if checkResult[0][0] > 100:
                        singleLockRecordList.append(homeAddr[0][0])  # 房间地址
                        singleLockRecordList.append(singleLock[12])  # mac地址
                        singleLockRecordList.append(singleLock[11])  # deviceID
                        singleLockRecordList.append(
                            (checkTime + "同步次数为:   " + str(checkResult[0][0]).decode("utf-8")))  # deviceID
                if singleLockRecordList:
                    searchResultList.append(singleLockRecordList)
    tmpTime = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    a = random.randint(0, 1000)
    with open(str(a) + "_tmp_" + tmpTime + 'test.txt', 'w') as f:
        for singleRecord in searchResultList:
            f.write(singleRecord[0] + "_" + singleRecord[1] + "_" + singleRecord[2] + "_" + singleRecord[3] + "\n")


class myThread(threading.Thread):  # 继承父类threading.Thread
    def __init__(self, threadID, name, contractInfoList):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
        self.contractInfoList = contractInfoList

    def run(self):  # 把要执行的代码写到run函数里面 线程在创建后会直接运行run函数
        print "开始 " + self.name
        checkLockSyncWithThreadsFuntion(self.threadID, "2018-05-04", self.contractInfoList)
        print "退出 " + self.name


def checkLockSyncWithThreads():
    getContractinfoSql = "SELECT DISTINCT contract_info.contractNum FROM mc_user,contract_info WHERE contract_info.`user`=mc_user.ID"
    getContractInfo = dbOperation(getContractinfoSql)

    div = len(getContractInfo) // 6
    mod = len(getContractInfo) % 3
    print div
    print(mod)

    c1 = getContractInfo[0:div]
    c2 = getContractInfo[div:div * 2]
    c3 = getContractInfo[div * 2:div * 3]
    c4 = getContractInfo[div * 3:div * 4]
    c5 = getContractInfo[div * 4:div * 5]
    c6 = getContractInfo[div * 5:len(getContractInfo) + 1]

    thread1 = myThread(1, "第一线程", c1)
    thread2 = myThread(2, "第二线程", c2)
    thread3 = myThread(3, "第三线程", c3)
    thread4 = myThread(4, "第四线程", c4)
    thread5 = myThread(5, "第五线程", c5)
    thread6 = myThread(6, "第六线程", c6)

    # 开启线程
    thread1.start()
    thread2.start()
    thread3.start()
    thread4.start()
    thread5.start()
    thread6.start()

    thread1.join()
    thread2.join()
    thread3.join()
    thread4.join()
    thread5.join()
    thread6.join()

    endtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    end = datetime.datetime.now()

    print u"开始计时: %s" % starttime
    print u"结束计时: %s" % endtime
    print u"总共花费时间为: %s" % (end - start)


# def transferMacToHouseAddress(username, password,address):
#     generateCenterControlInfoAtDeviceCenter()
#     host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
#     ck = getCookies(username, password)
#     pageCount = centerControlGetPageCount(username, password, address)
#     # 获取中控id
#     excelRowCount = 1
#     for index in range(pageCount):
#         pageNo = index + 1
#         # print "正在查询第%s页" % str(pageNo)
#         print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
#         payload = getCenterControlPayload(pageNo, address)
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]["result"]
#         rsp = rsp["resultList"]
#         for k in range(len(rsp)):
#             resultrow = rsp[k]
#             # ccid=centercontrolid
#             ccid = resultrow["id"]
#             # cc=centerControl
#             ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
#             ccMac = dbOperation(ccMacSql)
#             ccMac = ccMac[0][0]
#             print ccMac
#             houseAddress = resultrow["houseAddress"]
#             centerControlstatus = resultrow["status"]
#             centerCtrolVersion = resultrow["version"]
#
#             writeExcel(excelRowCount, 0, ccMac)
#             writeExcel(excelRowCount, 1, centerCtrolVersion)
#             if centerControlstatus == "1":
#                 writeExcel(excelRowCount, 2, u"离线", 2)
#             else:
#                 writeExcel(excelRowCount, 2, u"在线")
#             writeExcel(excelRowCount, 3, houseAddress)
#             excelRowCount = excelRowCount + 1
#     #         # # 通过 housei
#
#     print  "done"
#     a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
#     os.rename("test.xls", unicode(address[0], "utf-8") + u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')
# def delblankline(infile,outfile):
#     infopen = open(infile,'r')
#     outfopen = open(outfile,'w')
#     lines = infopen.readlines()
#     for line in lines:
#         if line.split():
#             outfopen.writelines(line)
#         else:
#             outfopen.writelines("")
#     infopen.close()
#     outfopen.close()
# def restartCenterControl(username, password):
#
#     # 读取txt，获取中控mac，获取id，然后重启
#     delblankline("ccinfo.txt","ccmac.txt")
#
#     with open("ccmac.txt", 'r') as f:
#         ccmacs = f.readlines()
#     ccmacList=[]
#     for a in ccmacs:
#
#         newMac=a[0:2]+":"+a[2:4]+":"+a[4:6]+":"+a[6:8]+":"+a[8:10]+":"+a[10:12]
#         # ccmacList.append(newMac)
#         getmacSql="SELECT id FROM center_control WHERE macAddress="+ "\'" +  newMac+ "\'" + ";"
#         getmacSqlResult=dbOperation(getmacSql)
#         print getmacSqlResult[0][0]
#         ccid = getmacSqlResult[0][0]
#
#
#         host = 'http://www.danbay.cn/system/socket/restartGate'
#         ck = getCookies(username, password)
#
#         payload = {"centerDeviceId": ccid}
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         print r.text
#
#         with open(u"test", 'a+') as f:
#             f.write(str(ccid)+"---"+r.text + '\n')
#
#     a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
#     os.rename("test", unicode("重启结果", "utf-8") + unicode(a, "utf-8") + u'.txt')

def restartCenterControlByAddr(username, password, address):
    host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
    ck = getCookies(username, password)
    pageCount = centerControlGetPageCount(username, password, address)
    for index in range(pageCount):
        pageNo = index + 1
        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        payload = getCenterControlPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]["result"]
        rsp = rsp["resultList"]
        for k in range(len(rsp)):
            resultrow = rsp[k]
            ccid = resultrow["id"]
            print ccid
            cchost = 'http://www.danbay.cn/system/socket/restartGate'
            ccpayload = {"centerDeviceId": ccid}
            ccr = requests.post(cchost, data=ccpayload, cookies=ck)
            ccrsp = json.loads(ccr.text)
            print ccr.text
            with open(u"test", 'a+') as f:
                f.write(str(ccid) + "---" + ccr.text + '\n')

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test", unicode("重启结果", "utf-8") + unicode(a, "utf-8") + u'.txt')


# def getOfflineCenterControl(username, password,address):
#     generateCenterControlExcel()
#     addedCenterID = []
#     host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#     ck = getCookies(username, password)
#     pageCount = getPageCount(username, password, address)
#     for index in range(pageCount):
#         pageNo=index+1
#         payload=getPayload(pageNo,address)
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]
#         rsp = rsp["resultList"]
#         for k in range(len(rsp)):
#             i = rsp[k]
#             houseid = i["hosueInfo"]["id"]
#             # # 通过 houseid 获取某个房间的设备
#             # req = "http://www.danbay.cn/system/house/getDeviceInfoByHouseId?id=%s" % houseid
#             # r = requests.get(req, cookies=ck)
#             # devicersp = r.text
#             devicersp=getJson(ck,houseid)
#
#             devicersp = json.loads(devicersp, encoding='utf-8')
#             devicersp = devicersp["result"]
#             gatewayList=devicersp["gatewayList"]
#             for deviceIndex in range(len(gatewayList)):
#                 gw=gatewayList[deviceIndex]
#                 homeaddress=getHomeAddress(houseid)
#                 gwID=gw["id"]
#                 if gwID not in addedCenterID:
#                     addedCenterID.append(gwID)
#                     print addedCenterID
#                     onlineDeviceNum=gw["onlineDeviceNum"]
#                     onlineStatus=gw["onlineStatus"]
#                     totalDeviceNum=gw["totalDeviceNum"]
#                     roomAddr=devicersp["roomName"]
#                     # print roomAddr
#                     #房间地址
#                     writeExcel(len(addedCenterID), 0, roomAddr)
#                     # 中控id
#                     writeExcel(len(addedCenterID), 1, gwID)
#                     # 在线状态
#                     if onlineStatus=="0":
#                         writeExcel(len(addedCenterID), 2, u"在线")
#                     elif onlineStatus=="1":
#                         writeExcel(len(addedCenterID), 2, u"离线",2)
#
#                     if onlineDeviceNum:
#                         writeExcel(len(addedCenterID), 3, onlineDeviceNum)
#                     if totalDeviceNum:
#                         writeExcel(len(addedCenterID), 4, totalDeviceNum)
#                     #中控版本，设备id，mac地址
#                     getCenterControlmacDevIDVersion="SELECT deviceId,macAddress,version FROM center_control WHERE id=%s"%gwID
#                     result=dbOperation(getCenterControlmacDevIDVersion)
#
#                     writeExcel(len(addedCenterID) ,5 , result[0][0])
#                     writeExcel(len(addedCenterID) ,6 , result[0][1])
#                     writeExcel(len(addedCenterID) ,7, result[0][2])
#                 # houseID，如果不为空就写进去
#
#                 #index*10+k*2+1+deviceIndex 行数    4 列数， meterID 需要写入的值
#                 # writeExcel(index*20+k*2+1+deviceIndex,4,meterID)
#
#     print  "done"
#     a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
#     os.rename("test.xls", unicode(address, "utf-8") + u'_中控在线离线状态_' + unicode(a, "utf-8") + u'.xls')
def getConf():
    cwd = os.getcwd()
    with open(cwd + "\info.txt", 'r') as f:
        accountInfo = f.readlines()
    address = accountInfo[0].strip().split("=")[1]
    address = address.split(",")
    account = {}
    account["address"] = address
    account["option"] = accountInfo[1].strip().split("=")[1]

    return account


# def getMeterCount(username, password, address):
#     meterExcel()
#     waterExcel()
#     host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
#     ck = getCookies(username, password)
#     pageCount = getPageCount(username, password, address)
#     meterRows=1
#     waterRows=1
#     for index in range(pageCount):
#         pageNo = index + 1
#         print "正在获取第%s页"%pageNo
#         payload = getPayload(pageNo, address)
#         r = requests.post(host, data=payload, cookies=ck)
#         rsp = json.loads(r.text)
#         rsp = rsp["result"]
#         rsp = rsp["resultList"]
#         for k in range(len(rsp)):
#             i = rsp[k]
#             houseid = i["hosueInfo"]["id"]
#             devicersp = getJson(ck, houseid)
#
#             devicersp = json.loads(devicersp, encoding='utf-8')
#             devicersp = devicersp["result"]
#             meterList = devicersp["meterList"]
#             for deviceIndex in range(len(meterList)):
#                 meter = meterList[deviceIndex]
#                 homeaddress = i["hosueInfo"]["homeAddress"]
#                 meterType = meter["meterType"]
#                 subType = meter["subType"]
#                 meterID = meter["deviceId"]
#                 elecmeterID = meter["id"]
#                 meterStatus = meter["onlineStatus"]
#                 aa = getHouseID(meterID)
#                 # 0 水表
#                 mydatelist=[21,20,19,18,17,16,15]
#                 myseconddatelist=[20,21,22,23,24,25,26]
#                 if meterType == "0":
#                     sql1 = "SELECT id,houseId from energy_device WHERE deviceId='" + meterID + "';"
#                     addrAndId = dbOperation(sql1)
#                     # roomaddress = addrAndId[0][1]
#                     roomaddress = homeaddress
#                     waterID = addrAndId[0][0]
#                     # sql = "SELECT * FROM energy_day_consumption WHERE energyDevice='" + str(
#                     #     waterID) + "' ORDER BY readTime DESC;"
#
#                     # 1月21号的值
#                     for mydate in mydatelist:
#                         w1sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
#                         waterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-01-"+str(mydate)+"';"
#                         firstwaterCount = dbOperation(w1sql)
#                         if firstwaterCount:
#                             fwaterCount= firstwaterCount[0][0]
#                             print w1sql,"水表1"
#                             break
#                     # 2月22或以后的值
#                     for myseconddate in myseconddatelist:
#                         w2sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
#                         waterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-02-"+str(myseconddate)+"';"
#                         secondwaterCount = dbOperation(w2sql)
#                         if secondwaterCount:
#                             SeCount= secondwaterCount[0][0]
#                             print w2sql,"水表2"
#                             break
#                     StotalWater=SeCount-fwaterCount
#                     print roomaddress,  StotalWater
#                     TwowaterWriteExcel(waterRows, 0, roomaddress)
#                     TwowaterWriteExcel(waterRows, 1, StotalWater)
#                     TwowaterWriteExcel(waterRows, 2, waterID)
#                     waterRows=waterRows+1
#
#                 elif meterType == "1":
#                     metersql1 = "SELECT id,houseId from energy_device WHERE deviceId='" + meterID + "';"
#                     meteraddrAndId = dbOperation(metersql1)
#                     # meterroomaddress = meteraddrAndId[0][1]
#                     meterID = meteraddrAndId[0][0]
#
#                     # 1月21号的值
#                     for myddate in mydatelist:
#                         d1sql= "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
#                             meterID) +"' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-01-"+str(myddate)+"';"
#                         # print  sql
#                         firstDianCount = dbOperation(d1sql)
#                         if firstDianCount:
#                             fDianCount= firstDianCount[0][0]
#                             print d1sql,"电表1"
#                             break
#
#                     # 2月22或以后的值
#                     for mysecondddate in myseconddatelist:
#                         d2sql = "SELECT meterCount FROM energy_day_consumption WHERE energyDevice='" + str(
#                             meterID) + "' and DATE_FORMAT(readTime,'%Y-%m-%d')='2018-02-" + str(
#                             mysecondddate) + "';"
#                         seconddianCount = dbOperation(d2sql)
#                         if seconddianCount:
#                             SedianCount = seconddianCount[0][0]
#                             print d2sql,"电表2"
#                             break
#
#                     DtotalDian=SedianCount-fDianCount
#                     TwometerWriteExcel(meterRows, 0, homeaddress,2)
#                     TwometerWriteExcel(meterRows,1,DtotalDian,2)
#                     TwometerWriteExcel(meterRows,2,meterID,2)
#                     # meterResult = dbOperation(metersql)
#                     # # print meterResult
#                     # if meterResult:
#                     #     meterCount = meterResult[0][2]
#                     #     meterreadTime = meterResult[0][4]
#                     #     # print meterroomaddress, meterreadTime, meterCount
#                     #     # meterWriteExcel()
#                     #     meterWriteExcel(meterRows,0,meterroomaddress)
#                     #     meterWriteExcel(meterRows,1,meterreadTime,2)
#                     #     meterWriteExcel(meterRows,2,meterCount)
#                     # else:
#                     #     meterWriteExcel(meterRows,0,meterroomaddress+"-"+metersql)
#
#                     meterRows=meterRows+1

# def getAddress(username, password):
#     addrSql="SELECT detailAddress FROM homesourcelist WHERE deleteState !=1 AND location LIKE '%深圳%' ;"
#     addrResult=dbOperation(addrSql,db='danbay_projects')
#     excelRowCount = 1
#     generateCenterControlInfoAtDeviceCenter()
#
#     # for dbaddr in addrResult:
#     #     if dbaddr[0]:
#     #         print dbaddr[0]
#     #         host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
#     #         ck = getCookies(username, password)
#     #         # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
#     #         pageCount = centerControlGetPageCount(username, password, dbaddr[0])
#     #         # 获取中控id
#     #         if pageCount !=0:
#     #             for index in range(pageCount):
#     #                 pageNo = index + 1
#     #                 print "正在查询第%s页" % str(pageNo)
#     #                 payload = getCenterControlPayload(pageNo, dbaddr[0])
#     #                 r = requests.post(host, data=payload, cookies=ck)
#     #                 rsp = json.loads(r.text)
#     #                 rsp = rsp["result"]["result"]
#     #                 rsp = rsp["resultList"]
#     #                 for k in range(len(rsp)):
#     #                     resultrow = rsp[k]
#     #                     ccid = resultrow["id"]
#     #                     # cc=centerControl
#     #                     ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
#     #                     ccMac = dbOperation(ccMacSql)
#     #                     ccMac = ccMac[0][0]
#     #                     print ccMac
#     #                     houseAddress = resultrow["houseAddress"]
#     #                     centerControlstatus = resultrow["status"]
#     #                     centerCtrolVersion = resultrow["version"]
#     #                     print houseAddress,centerControlstatus,centerCtrolVersion
#     #
#     #                     writeExcel(excelRowCount, 0, ccMac)
#     #                     writeExcel(excelRowCount, 1, centerCtrolVersion)
#     #                     if centerControlstatus == "1":
#     #                         writeExcel(excelRowCount, 2, u"离线", 2)
#     #                     else:
#     #                         writeExcel(excelRowCount, 2, u"在线")
#     #                     writeExcel(excelRowCount, 3, houseAddress)
#     #                     excelRowCount = excelRowCount + 1
#     #                     print excelRowCount, "Excel 行数"
#
#
#
#     # 获取所有项目的中控在线离线状态
#     excelRowCount = 1
#     generateCenterControlInfoAtDeviceCenter()
#     host = 'http://www.danbay.cn/system/centerControl/findCenterControlSetting'
#     ck = getCookies(username, password)
#     # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
#     # pageCount = centerControlGetPageCount(username, password, dbaddr[0])
#     payload = getCenterControlPayload("1", "深圳")
#     host = "http://www.danbay.cn/system/centerControl/findCenterControlSetting"
#     ck = getCookies(username, password)
#     r = requests.post(host, data=payload, cookies=ck)
#     rsp = json.loads(r.text)
#     rsp = rsp["result"]["result"]
#     pageCount = rsp["pageCount"]
#
#     print "总共的页数是:", pageCount
#     # return pageCount
#     # 获取中控id
#     if pageCount != 0:
#         for index in range(pageCount):
#             pageNo = index + 1
#             print "正在查询第%s页" % str(pageNo)
#             # payload = getCenterControlPayload(pageNo, dbaddr[0])
#             payload = {'pageNo': pageNo, 'pageSize': '8', "status": "", 'likeStr': "深圳", "isNewVersion": ""}
#             # return payload
#             r = requests.post(host, data=payload, cookies=ck)
#             rsp = json.loads(r.text)
#             rsp = rsp["result"]["result"]
#             rsp = rsp["resultList"]
#             for k in range(len(rsp)):
#                 resultrow = rsp[k]
#                 ccid = resultrow["id"]
#                 # cc=centerControl
#                 ccMacSql = "SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"
#                 ccMac = dbOperation(ccMacSql)
#                 ccMac = ccMac[0][0]
#                 print ccMac
#                 houseAddress = resultrow["houseAddress"]
#                 centerControlstatus = resultrow["status"]
#                 centerCtrolVersion = resultrow["version"]
#                 print houseAddress, centerControlstatus, centerCtrolVersion
#
#                 writeExcel(excelRowCount, 0, ccMac)
#                 writeExcel(excelRowCount, 1, centerCtrolVersion)
#                 if centerControlstatus == "1":
#                     writeExcel(excelRowCount, 2, u"离线", 2)
#                 else:
#                     writeExcel(excelRowCount, 2, u"在线")
#                 writeExcel(excelRowCount, 3, houseAddress)
#                 excelRowCount = excelRowCount + 1
#                 print excelRowCount, "Excel 行数"
#     # print  "done"
#     # a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
#     # os.rename("test.xls",  u'_获取中控信息表_' + unicode(a, "utf-8") + u'.xls')
def delLockPwd(username, password, address):
    generateLockPwdCount()
    lockPwdCountRow = 1
    payload = {'pageNo': '1', 'pageSize': '10', 'likeStr': address, "detailAddress": "", "floor": "", "spaceType": "",
               "userId": "1"}
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    r = requests.post(host, data=payload, cookies=ck)
    rsp = json.loads(r.text)
    rsp = rsp["result"]
    pageCount = rsp["pageCount"]
    for index in range(pageCount):
        pageNo = index + 1
        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        payload = getPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        for roomRecord in range(len(rsp)):
            i = rsp[roomRecord]
            # print i
            homeaddress = i["hosueInfo"]["homeAddress"]

            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            locklist = devicersp["lockList"]
            lockRoomName = devicersp["roomName"]
            for deviceIndex in range(len(locklist)):
                lock = locklist[deviceIndex]
                lcokDeviceId = lock["deviceId"]
                # print lcokDeviceId
                if lcokDeviceId:
                    print homeaddress
                    LockRecordPwdhost = 'http://www.danbay.cn/system/devicePwdInfo/getPwdByDeviceId'
                    LockRecordPwdpayload = {"pageNo": "1", "pageSize": "40", "deviceId": lcokDeviceId, "pwdType": "3"}
                    ck = getCookies(username, password)
                    LockRecordPwdRsp = requests.post(LockRecordPwdhost, data=LockRecordPwdpayload, cookies=ck)
                    LockRecordPwdRsp = json.loads(LockRecordPwdRsp.text)
                    LockRecordPwdRsp = LockRecordPwdRsp["result"]
                    LockRecordPwdRsp = LockRecordPwdRsp["list"]
                    LockRecordPwdRsp = LockRecordPwdRsp["resultList"]

                    # 只查找租客密码记录数
                    writeExcel(lockPwdCountRow, 0, homeaddress)
                    writeExcel(lockPwdCountRow, 1, len(LockRecordPwdRsp))
                    writeExcel(lockPwdCountRow, 2, lcokDeviceId)
                    lockPwdCountRow = lockPwdCountRow + 1

                    # 查找没有开门记录的租客密码记录数
                    # noOpenRecord=0
                    # for pwdRecord in LockRecordPwdRsp:
                    #     print pwdRecord["id"]
                    #     # 通过密码ID找出开门记录
                    #     getLockPwdCountSql="SELECT * FROM log_report WHERE device_id="+ "\'" +str(lcokDeviceId)+"' AND  url_path LIKE '%status/open%' AND report_content LIKE '%"+str(pwdRecord["id"])+"%'  ORDER BY res_time DESC"
                    #     getLockPwdCountSqlResult=dbOperation(getLockPwdCountSql,db='danbay_task')
                    #     if not getLockPwdCountSqlResult:
                    #         print "无开门记录"
                    #         noOpenRecord=noOpenRecord+1
                    # if   noOpenRecord>5:
                    #     writeExcel(lockPwdCountRow, 0, homeaddress)
                    #     writeExcel(lockPwdCountRow, 1, noOpenRecord)
                    #     writeExcel(lockPwdCountRow, 2, lcokDeviceId)
                    #
                    #     lockPwdCountRow = lockPwdCountRow + 1
                    noOpenRecord = 0
                    # 开始删除密码
                    # 删除门锁密码
                    # deviceId
                    # pwdType
                    # pwdID
                    # mtoken
                    # 先获取密码id

    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))
    os.rename("test.xls", unicode(address, "utf-8") + u'_门锁租客密码数统计_' + unicode(a, "utf-8") + u'.xls')


def getSheBeiDushu(username, password, address):
    excelRows = 1
    dushuExcel()
    rb = xlrd.open_workbook("test.xls", formatting_info=True)
    wb = copy(rb)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'yyyy-mm-dd hh:mm:ss'

    # 通过 地址 过滤出房源信息
    host = 'http://www.danbay.cn/system/house/getHouseInfoByCondition'
    ck = getCookies(username, password)
    pageCount = getPageCount(username, password, address)
    # pageCount 为根据房源地址过滤出来的所有房间地址信息，就是某个项目下的所有信息
    # pageCount=46
    for index in range(pageCount):
        pageNo = index + 1
        payload = getPayload(pageNo, address)
        r = requests.post(host, data=payload, cookies=ck)
        rsp = json.loads(r.text)
        rsp = rsp["result"]
        rsp = rsp["resultList"]
        print "当前页数是第%s页,总共页数是%s页".decode("utf-8") % (str(pageNo), str(pageCount))
        # 判断是否有水电表

        for k in range(len(rsp)):
            i = rsp[k]
            houseid = i["hosueInfo"]["id"]
            devicersp = getJson(ck, houseid)
            devicersp = json.loads(devicersp, encoding='utf-8')
            devicersp = devicersp["result"]
            meterList = devicersp["meterList"]
            if devicersp["meterList"]:
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

                    # writeExcel(excelRows, 0, homeaddress)
                    wb.get_sheet(0).write(excelRows, 0, homeaddress)
                    # 设备类型，水保还是电表

                    if meterType == "0":
                        if subType:
                            if subType == "1":
                                # writeExcel(excelRows, 1, u"热水")
                                wb.get_sheet(0).write(excelRows, 1, u"热水")
                            elif subType == "0":
                                # writeExcel(excelRows, 1, u"冷水")
                                wb.get_sheet(0).write(excelRows, 1, u"冷水")
                        else:
                            # writeExcel(excelRows, 1, u"水表")
                            wb.get_sheet(0).write(excelRows, 1, u"水表")
                    elif meterType == "1":
                        # writeExcel(excelRows, 1, u"电表")
                        wb.get_sheet(0).write(excelRows, 1, u"电表")
                    # 在线状态，离线还是在线
                    if meterStatus == "1":
                        # writeExcel(excelRows, 2, u"离线",2)
                        wb.get_sheet(0).write(excelRows, 2, u"离线", redStyle())
                        # 写上离线时间，计算离线了几天

                    elif meterStatus == "0":
                        # writeExcel(excelRows, 2, u"在线")
                        wb.get_sheet(0).write(excelRows, 2, u"在线")
                    # houseID，如果不为空就写进去
                    # 获取框读数
                    kuangDushu = meter["meterCount"]
                    print kuangDushu
                    print homeaddress
                    wb.get_sheet(0).write(excelRows, 3, kuangDushu)
                    energyHost = "http://www.danbay.cn/system/engeryDevice/getEnergyDeviceByHouseInfoId"
                    detailpayload = {"houseInfoId": str(houseid), "type": "1", "meterId": str(elecmeterID)}
                    detailResopen = requests.post(energyHost, data=detailpayload, cookies=ck)
                    detailResopen = json.loads(detailResopen.text)
                    detailResopen = detailResopen["result"]
                    detailResopen = detailResopen["resultList"]
                    realDu = detailResopen[0]["real_meter_data"]  # 默认只取第一个电表的读数，默认只有一个电表
                    wb.get_sheet(0).write(excelRows, 4, realDu)

                    # #根据deviceid找到对应的中控，并写入进去
                    # getmeterControlIdSql="SELECT centerControlId,address FROM energy_device WHERE deviceId="+ "\'" + meterID + "\'" + ";"
                    #
                    # centerCotrolID=dbOperation(getmeterControlIdSql)[0][0]
                    # meterMac=dbOperation(getmeterControlIdSql)[0][1]
                    # # writeExcel(excelRows, 3, meterMac)
                    # wb.get_sheet(0).write(excelRows, 3, meterMac)
                    # # writeExcel(excelRows, 4, meterID)
                    # wb.get_sheet(0).write(excelRows, 4, meterID)
                    #
                    # getmeterControlAddress="SELECT address FROM center_control WHERE id="+ "\'" + str(centerCotrolID) + "\'" + ";"
                    # meterControlAddress=dbOperation(getmeterControlAddress)
                    # # 水电表关联的中控的地址信息
                    #
                    # # writeExcel(excelRows, 5, meterControlAddress[0][0])
                    # wb.get_sheet(0).write(excelRows, 5, meterControlAddress[0][0])
                    # #中控deviceid
                    # ControlAddress = "SELECT deviceId,macAddress,online FROM center_control WHERE id=" + "\'" + str(centerCotrolID) + "\'" + ";"
                    # centID=dbOperation(ControlAddress)
                    # # print centID[0][0],centID[0][1]
                    #
                    # # writeExcel(excelRows, 6, centID[0][1])
                    # wb.get_sheet(0).write(excelRows, 6, centID[0][1])
                    # # 中控在线离线状态
                    # if centID[0][2]==0:
                    #      # writeExcel(excelRows, 7, u'在线')
                    #      wb.get_sheet(0).write(excelRows, 7, u'在线')
                    # elif centID[0][2]==1:
                    #      # writeExcel(excelRows, 7, u'离线',2)
                    #     wb.get_sheet(0).write(excelRows, 7, u'离线',redStyle())
                    #
                    # # writeExcel(excelRows, 8, centID[0][0])
                    # wb.get_sheet(0).write(excelRows, 8, centID[0][0])
                    # print "完成房间:".decode("utf-8"),homeaddress,"的设备信息录入...".decode("utf-8")
                    excelRows = excelRows + 1
        wb.save(u'test.xls')
    a = time.strftime('%Y-%m-%d_%H_%M_%S', time.localtime(time.time()))

    os.rename("test.xls", unicode(address, "utf-8") + u'_水电表状态_' + unicode(a, "utf-8") + u'.xls')


def checkDeviceOffLine(deviceID):
    # 计算一个设备一个月内的离线时长
    # 一个设备离线率=离线总时长除以一个月的时间
    # 计算30天之内的离线时长
    # deviceID="129286e6a7fc9fec16b3da416d785250"
    checkStartTime=(datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
    checkEndTime=datetime.datetime.now().strftime("%Y-%m-%d")

    checkSql = "SELECT url_path,report_time from log_report WHERE device_id='" + deviceID + "' and url_path LIKE '%log%' AND report_time BETWEEN '"+checkStartTime+" 00:00:00' AND  '"+checkEndTime+" 00:00:00' ORDER BY res_time ASC;"
    checkResult = dbOperation(checkSql, db="danbay_task")
    offlineDic = {}
    offlineDicIndex = 0
    gloabolCheck = 0
    for i in range(len(checkResult)):
        recordList = []
        if gloabolCheck != 0:
            gloabolCheck = i
        try:
            if (i+1)!=len(checkResult):
                if "logout" in checkResult[i][0] and "login" in checkResult[i + 1][0]:
                    recordList.append(checkResult[i])
                    recordList.append(checkResult[i + 1])
                    offlineDic[offlineDicIndex] = recordList
                    offlineDicIndex = offlineDicIndex + 1
                    i = i + 2
                    gloabolCheck = i
                    # 这种情况是：假设checkResult的长度是7，当i=5，是有logout，然后i=6 有login，那这个时候刚刚好配对完

                    if i >= len(checkResult):
                        break
                else:
                    i = gloabolCheck
        except:
            print checkResult
            print u"checkDeviceOffLine有问题，程序出错了！！！"

    # print offlineDic
    timList = []
    for v in offlineDic.itervalues():
        # print v #v=[(u'device/logout', datetime.datetime(2018, 4, 1, 0, 45, 2)), (u'device/login', datetime.datetime(2018, 4, 1, 0, 45, 11))]
        # print (v[1][1] -v[0][1])
        timList.append((v[1][1] - v[0][1]))
    # print timList
    sunmT = datetime.timedelta(0, 1)
    for singleTIme in timList:
        sunmT = sunmT + singleTIme
    return sunmT


# def getAllCenterControlInfo():
#     generateAllCenterControlInfoFromDataBase()
#     ccdbSql="SELECT macAddress FROM center_control WHERE id=" + "\'" + str(ccid) + "\'" + ";"

if __name__ == '__main__':
    starttime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    start = datetime.datetime.now()
    print "Start Time: %s" % starttime

    config = getConf()
    config["username"] = "idjjd"
    config["password"] = "dknjghsd@"

    print u"过滤地址为：", config["address"][0].decode("utf-8")
    if config["option"] == str(1):
        print u"你所操作的选项是:", config["option"], u",获取门锁密码总数"
        getLockPwdCount(config["username"], config["password"], config["address"])  # 获取门锁预置密码以及正式密码的总数
    elif config["option"] == str(2):
        print u"你所操作的选项是:", config["option"], u",获取门锁在线离线状态"
        LockOnlineStatus(config["username"], config["password"], config["address"])  # 查看门锁在线离线状态
    elif config["option"] == str(3):
        # 根据地址循环获取
        checkLockDataSync(config["username"], config["password"], config["address"][0],10,1)
        # checkAllLockSyncsInDB()
        # checkLockSyncWithThreads()
        # checkDeviceOffLine()
        # checkDeviceOffline(config["username"], config["password"], config["address"][0], 1, 1)

    elif config["option"] == str(4):
        print u"你所操作的选项是:", config["option"], u",获取水电表在线离线读数状态"
        getAmmeterDeviceId(config["username"], config["password"], config["address"][0])  # 获取水电表在线离线状态 以及 对应中控在线离线状态
    elif config["option"] == str(5):
        print u"你所操作的选项是:", config["option"], u",获取中控状态,版本信息"
        getCenterControlInfoByDeviceCenter(config["username"], config["password"],
                                           config["address"])  # 在设备中心获取中控的信息在线离线，以及 mac信息
    elif config["option"] == str(6):
        print u"你所操作的选项是:", config["option"], u",批量重启中控"
        restartCenterControlByAddr(config["username"], config["password"], config["address"][0])
    elif config["option"] == "a":
        #分散式门锁在线离线状态
        fensanLockOnlineStatus(config["username"], config["password"])
    elif config["option"] == "b":
        #碧桂园门锁密码统计
        # BGYLockPwd(config["username"], config["password"])
        getFenSanCenterControlInfo(config["username"], config["password"])




    # restartCenterControl(config["username"], config["password"])

    # delLockPwd(config["username"], config["password"], "广东省深圳市罗湖区桂园北街92号11栋")
    # getShuiDianInfoFromDB("广东省深圳市龙岗区广州优家资产管理有限公司")
    # getSheBeiDushu(config["username"], config["password"],"稻香路")

    endtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    end = datetime.datetime.now()

    print u"开始计时: %s" % starttime
    print u"结束计时: %s" % endtime
    print u"总共花费时间为: %s" % (end - start)
    print "\n"
    print u"按任意键退出"
    raw_input("")
    os._exit(0)
