#!/usr/bin/env python
import sys
import os.path
import time
# import csv
import pandas as pd
import pymongo
# import json
import plistlib
from datetime import datetime
import threading
import re
import hashlib
import logging

g_rawCount = 0

'''
def import_content(filepath,pTable):
    data = pd.read_csv(filepath)
    data_json = json.loads(data.to_json(orient='records'))
    pTable.remove()
    pTable.insert(data_json)
'''


def import_content(filepath, pTable):
    with open(filepath) as pfile:
        lines_ = pfile.readlines()
        strHeader = lines_[0].replace(".","-")
        newHeader = strHeader.split(',')
        pTable.remove()
        try:
            for i1 in range(1, len(lines_)):
                dicValues = dict(zip(newHeader, lines_[i1].split(',')))
                pTable.insert(dicValues)
        except Exception as e5:
            logging.error("import_contenet : %s" % e5)
    pfile.close()


def is_date(pInputStr):
    try:
        pd.to_datetime(pInputStr)
        return True
    except Exception as e6:
        print ("error : %s" % e6)
        return False


def find_Duplicate(pHeader, pVersion, pTable):
    bSkipAlisa = False
    bSkipLimit = False
    if pHeader in g_dictAlias.keys():
        bSkipAlisa = True
    if pVersion in g_LimitData and bSkipAlisa is True:
        bSkipLimit = True
    return bSkipAlisa, bSkipLimit


def TransferFieldName(pHeader, pMD5_Header):
    for item in pHeader:
        if item in g_ReserverdNames:
            pMD5_Header.append(item)
        else:
            md5_hex = hashlib.md5(item.encode("UTF-8")).hexdigest()
            pMD5_Header.append(md5_hex)


def Thread_Import_Limit(pTable, pVersion, pHeader, pPriority, pUpperLimit, pLowerLimit, pUnit, pData, pRange, pCreated):
    start_time = datetime.now()
    logging.info("limit range: {0} - {1}".format(pRange[0], pRange[-1]))
    for idx in pRange:
        dicValues = dict()
        dicValues['Product'] = pData[0]
        dicValues['Version'] = pVersion
        dicValues['OriginalName'] = pHeader[idx]
        dicValues['DisplayName'] = ""
        dicValues['FieldType'] = "string"
        if len(pPriority) > idx:
            dicValues['PDCAPriority'] = pPriority[idx]
        else:
            dicValues['PDCAPriority'] = ""
        if len(pUpperLimit) > idx:
            dicValues['UpperLimit'] = pUpperLimit[idx]
            if pUpperLimit[idx].isdigit():
                dicValues['FieldType'] = 'float'
        else:
            dicValues['UpperLimit'] = ""
        if len(pLowerLimit) > idx:
            dicValues['LowerLimit'] = pLowerLimit[idx]
            if pLowerLimit[idx].isdigit():
                dicValues['FieldType'] = 'float'
        else:
            dicValues['LowerLimit'] = ""
        if len(pUnit) > idx:
            dicValues['MeasurementUnit'] = pUnit[idx]
        else:
            dicValues['MeasurementUnit'] = ""
        dicValues['ctime'] = datetime.utcnow()
        if len(pData) > idx:
            if pData[idx].isdigit():
                dicValues['FieldType'] = 'float'
            if is_date(pData[idx]):
                dicValues['FieldType'] = 'DateTime'
        try:
            #if dicValues['FieldType'] == 'float':
            pTable.insert(dicValues)
            if pCreated is True:
                pTable.create_index([("Version", pymongo.ASCENDING), ("OriginalName", pymongo.ASCENDING)], unique=True)
                pCreated = False
        except pymongo.errors.DuplicateKeyError as e1:
            print ("error : %s" % e1)
        except Exception as e2:
            logging.error("%s" % e2)
            global g_Exception
            g_Exception = True
    end_time = datetime.now()
    logging.info("{0} - Duration time for import limit : {1}".format(pRange[0], end_time - start_time))


def Thread_Import_Alisa(pOldHeader, pnewHeader, pTable, pRange, pCreated):
    start_time = datetime.now()
    logging.info("Alisa range: {0} - {1}".format(pRange[0], pRange[-1]))
    for idx in pRange:
        dicValues = dict()
        dicValues['originalname'] = pOldHeader[idx]
        dicValues['definer'] = 'Upload_Agent'
        dicValues['ctime'] = datetime.utcnow()
        dicValues['redefinename'] = pnewHeader[idx]
        try:
            pTable.insert(dicValues)
            if pCreated is True:
                pTable.create_index([("originalname", pymongo.ASCENDING)], unique=True)
                pCreated = False
        except pymongo.errors.DuplicateKeyError as e3:
            print ("error : %s" % e3)
        except Exception as e4:
            logging.error("%s" % e4)
            global g_Exception
            g_Exception = True
    end_time = datetime.now()
    logging.info("{0} - Duration time for import Alias : {1}".format(pRange[0], end_time - start_time))


def import_Limit_Alias(pContent, pnewHeader, pTable1, pTable2):
    oldheader = pContent[1].split(',')
    data = pContent[7].split(',')
    tmpVersion = pContent[0].split('Version:')
    version = tmpVersion[1].split(',')
    Priority = pContent[3].split(',')
    UpperLimit = pContent[4].split(',')
    LowerLimit = pContent[5].split(',')
    Unit = pContent[6].split(',')
    for idx in range(0, len(version)):
        version[idx] = version[idx].strip()
    AlisaTotal = pTable1.count()
    LimitTotal = pTable2.count()
    md5_oldheader = hashlib.md5(pContent[1].encode("UTF-8")).hexdigest()
    bSkipAlisa, bSkipLimit = find_Duplicate(md5_oldheader, version[0], pTable1)
    if AlisaTotal == 0:
        bAlisaCreated = True
    else:
        bAlisaCreated = False
    if bSkipAlisa is False:
        g_ThreadArray.append(threading.Thread(target=Thread_Import_Alisa,
                                          args=(oldheader, pnewHeader, pTable1, range(0, len(oldheader)), bAlisaCreated)))
        for t3 in g_ThreadArray:
            t3.start()
    if LimitTotal == 0:
        bLimitCreated = True
    else:
        bLimitCreated = False
    if bSkipLimit is False:
        Thread_Import_Limit(pTable2, version[0], oldheader, Priority, UpperLimit, LowerLimit, \
                            Unit, data, range(0, len(oldheader)), bLimitCreated)
    g_LimitData.append(version[0])
    g_dictAlias.update({md5_oldheader: pnewHeader})


def Thread_import_PDCA(plines, pHeader, pTable, pRange, pCreated, pLock):
    global g_rawCount
    start_time = datetime.now()
    logging.info("PDCA range: {0} - {1}".format(pRange[0], pRange[-1]))
    for idx in pRange:
        rx = '[' + re.escape(''.join(['\n', '\r'])) + ']'
        newline = re.sub(rx, '', plines[idx])
        newHeader = list()
        newData = list()
        newlist = newline.split(',')
        for idx1 in range(0, len(pHeader)):
            if pHeader[idx1] in g_ReserverdNames:
                newHeader.append(pHeader[idx1])
                newData.append(newlist[idx1])
            else:
                newHeader.append(pHeader[idx1])
                newData.append(newlist[idx1])
        dicValues = dict(zip(newHeader, newData))
        try:
            bMatch = False
            for item in g_LinesArray:
                if re.match(item, dicValues["Station ID"]):
                    bMatch = True
                    break
            if bMatch is True:
                pTable.insert(dicValues)
                pLock.acquire()
                g_rawCount += 1
                pLock.release()
                if pCreated is True:
                    pTable.create_index([("SerialNumber", pymongo.ASCENDING), ("StartTime", pymongo.ASCENDING)],
                                        unique=True)
                    pCreated = False
        except pymongo.errors.DuplicateKeyError as e7:
            print ("error : %s" % e7)
        except Exception as e8:
            logging.error("%s" % e8)
            global g_Exception
            g_Exception = True
            break
    end_time = datetime.now()
    logging.info("{0} - Duration time for import PDCA : {1}".format(pRange[0], end_time - start_time))


def import_PDCA_Data(pfile, pTable1, pTable2, pTable3):
    global g_rawCount
    content = pfile.readlines()
    nCount = len(content)
    if nCount >= 8:
        rx = '[' + re.escape(''.join(['\n', '\r'])) + ']'
        for idx in range(0, 8):
            content[idx] = re.sub(rx, '', content[idx])
        oldheader = content[1].split(',')
        newHeader = list()
        TransferFieldName(oldheader, newHeader)
        g_rawCount = 0
        lock = threading.Lock()
        count = 7
        nStep = int((nCount - count) / 3)
        ThreadList = list()
        if nStep <= 1:
            nStep = len(content)
        while True:
            new_count = count + nStep
            if new_count >= len(content):
                pRange = range(count, len(content))
                new_count = -1
            else:
                pRange = range(count, new_count)
            if count == 7 and pTable1.count() == 0:
                bCreated = True
            else:
                bCreated = False
            ThreadList.append(threading.Thread(target=Thread_import_PDCA,
                                               args=(content, newHeader, pTable1, pRange, bCreated, lock)))
            count = count + nStep
            if new_count == -1:
                break
        for t1 in ThreadList:
            t1.start()
        for t2 in ThreadList:
            t2.join()
        logging.info("Total import {0} raw data".format(g_rawCount))
        if g_rawCount > 0:
            import_Limit_Alias(content, newHeader, pTable2, pTable3)
    return nCount


Tool_Ver = 2.1
logging.basicConfig(format='%(asctime)-15s %(levelname)s => %(message)s', filename='Data_Uploader.log',
                    level=logging.DEBUG)
logging.info("Tool Ver. = %s" % Tool_Ver)
PlistName = "Data_Uploader.plist"
pl = plistlib.readPlist(PlistName)
Db_Url = pl["MongoDB_URL"]
collect = pymongo.uri_parser.parse_uri(Db_Url)
varCutTime = pl["CutTime"]
g_ReserverdNames = pl["ReserverdItemsName"]
try:
    mng_client = pymongo.MongoClient(collect['nodelist'][0][0], collect['nodelist'][0][1])
    mng_client.admin.authenticate(collect['username'], collect['password'])
except pymongo.errors.ConnectionFailure as e9:
    logging.error("Connection Failure %s" % e9)
    sys.exit(-1)
mng_db = mng_client['ndcweb']
Ary_Matrix = pl["Matrix"]
for varItem in Ary_Matrix:
    varPath = varItem["Path"]
    varTblName = varItem["TableName"]
    varFTime = varItem["TransferTime"]
    db_cm = mng_db[varTblName]
    mtTime = os.stat(varPath).st_mtime
    diffTime = time.time() - mtTime
    if diffTime >= varCutTime and mtTime != varFTime:
        szTime = time.strftime("{%Y-%m-%d %H:%M:%S}", time.localtime(varFTime))
        logging.info("Start Transfer Matrix : {0} - {1}".format(varPath, szTime))
        import_content(varPath, db_cm)
        varItem["TransferTime"] = mtTime
        plistlib.writePlist(pl, PlistName)
        szTime = time.strftime("{%Y-%m-%d %H:%M:%S}", time.localtime(mtTime))
        logging.info("Transfer Matrix Ending: {0} - {1}".format(varPath, szTime))
Ary_PDCA = pl["PDCA"]
g_LinesArray = pl["LineManager"]
for varItem in Ary_PDCA:
    path = varItem["Path"]
    varTblName = varItem["TableName"]
    varFTime = varItem["TransferTime"]
    varLimintTbl = varItem["Limit"]
    varAliasTbl = varItem["Alias"]
    g_dictAlias = dict()
    g_LimitData = list()
    db_cm = mng_db[varTblName]
    db_Alias = mng_db[varAliasTbl]
    db_Limit = mng_db[varLimintTbl]
    if not os.path.exists(path):
        continue
    for name in os.listdir(path):
        fullname = os.path.join(path, name)
        if os.path.isfile(fullname):
            mtTime = os.stat(fullname).st_mtime
            diffTime = time.time() - mtTime
            if diffTime >= varCutTime and mtTime != varFTime:
                g_ThreadArray = list()
                g_Exception = False
                logging.info("Start Transfer PDCA : {0} - {1}".format(fullname, os.stat(fullname).st_size))
                with open(fullname) as f:
                    RecCount = import_PDCA_Data(f, db_cm, db_Alias, db_Limit)
                f.close()
                varItem["TransferTime"] = mtTime
                plistlib.writePlist(pl, PlistName)
                for t in g_ThreadArray:
                    t.join()
                if g_Exception is False:
                    os.remove(fullname)
                szTime = time.strftime("{%Y-%m-%d %H:%M:%S}", time.localtime(mtTime))
                logging.info("Transfer PDCA Ending: {0} - {1}".format(RecCount, szTime))
'''
'''
mng_client.close()
logging.info("Finished")
