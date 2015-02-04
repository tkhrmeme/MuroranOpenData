#!/usr/local/bin/python
# -*- coding: utf-8 -*-

'''
道南バスの時刻表データのExcelファイルを読み込んで
routes, trips, service, stop_times
出力する。

出力データフォーマット：
    1.routes.txt
        route_type=3 (Bus)
        route_id: シートの名称（番号）
        agency_id: "Dounan-Bus"
        route_short_name: シート２列目の系統番号。空文字列の場合が有る。
        route_long_name: ファイル名の()に挟まれた名称
        route_desc: ルートの説明。AからBを経由してCまで

    2.trips.txt
        route_id: シートの名称（番号）
        service_id: Workday / Holiday 
        trip_id: シートの名称（番号）+ シーケンス番号
        trip_headsign: 行先表示
        trip_short_name: 系統番号
        direction_id: 上り 0，下り 1
        

    3.stop_times.txt
        trip_id: シートの名称（番号）+ シーケンス番号
        arrival_time: HH:MM:SS
        departure_time: HH:MM:SS
        stop_id: 各停留所のID
        stop_sequence: n++
        
    4.service.txt
        service_id: Workday / Holiday
        start_date
        end_date
        monday
        tuesday
        wednesday
        thursday
        friday
        saturday
        sunday
'''

import sys
import os
import xlrd
import csv

# Excelファイルのレイアウト
ROW_TOP = 0
ROW_HEADER = 1
ROW_DAYTYPE = 2
ROW_DATA_START = 3

COL_BUSSTOPS = 0
COL_KEITOU = 1
COL_DATA_START = 2

# 運行曜日
SERVICE_WORKDAY = 'workday_0' # 平日
SERVICE_HOLIDAY = 'holiday_0' # 土日祝

# stops.txtを読み込んで辞書（名称→ID）を作る
busStops = {}

# 各ファイルのポインタ
fp_routes = None
fp_trips = None
fp_times = None

agencyId = 'Hokkaido_Dounan_Bus'

routesHeader = "agency_id,route_type,route_id,route_short_name,route_long_name,route_text_color,route_color,route_url"
tripsHeader = "route_id, service_id, trip_id, trip_headsign, trip_short_name, direction_id"
timesHeader = "trip_id, arrival_time, departure_time, stop_id, stop_sequence"

## 停留所の名称から id を取得
def findBusStop(name):
    # TODO:辞書にして速くする
    if name in busStops:
        return busStops[name]
    return None

## ExcelをGTFS (CSV)に変換してファイルに出力
def convertExcelFile(filepath):

    (path, filename) = os.path.split(filepath)
    
    filename = filename[:-4]
    # 交通手段：バス
    routeType = 3
    # 経路の名称
    routeLongName = filename[filename.find('(')+1:filename.find(')')]
    
    try:
        filename = "routes.txt"
        newfile = os.path.join(path, filename)
        fp_routes = open(newfile,'w')
        fp_routes.write(routesHeader)
        fp_routes.write('\n')

        filename = "trips.txt"
        newfile = os.path.join(path, filename)
        fp_trips = open(newfile,'w')
        fp_trips.write(tripsHeader)
        fp_trips.write('\n')

        filename = "stop_times.txt"
        newfile = os.path.join(path, filename)
        fp_times = open(newfile,'w')
        fp_times.write(timesHeader)
        fp_times.write('\n')
    except IOError:
        print("File open error.")
        return
    
    ## Excelデータ読み取り
    book = xlrd.open_workbook(filepath, formatting_info=True)    
    for sheet_name in book.sheet_names():
        convertExcelSheet(book, routeLongName, sheet_name, fp_routes, fp_trips, fp_times)
        
    fp_routes.close()
    fp_trips.close()
    fp_times.close()

def outputRoute(fp, routeType, routeId, routeShortName, routeLongName):
    
    str = "{agency_id},{route_type},{route_id},{short_name},{long_name},{url},{text_color},{color}\n".format(
            agency_id=agencyId,
            route_type=routeType,
            route_id=routeId,
            short_name=routeShortName,
            long_name=routeLongName,
            url='',
            text_color='',
            color='' )
            
    fp.write(str)
    
def	outputTrip(fp, trip):
   
    str = "{routeId},{serviceId},{tripId},{headsign},{shortName},{direction}\n".format(
            routeId=trip['route_id'],
            serviceId=trip['service_id'],
            tripId=trip['trip_id'],
            headsign=trip['headsign'],
            shortName=trip['shortName'],
            direction=trip['direction']
            )
    fp.write(str)

def outputStopTimes(fp, stopTimes):
    
    str = "{tripId},{arrival},{departure},{stopId},{sequence}\n".format(
            tripId=stopTimes['trip_id'],
            arrival=stopTimes['arrival_time'],
            departure=stopTimes['departure_time'],
            stopId=stopTimes['stop_id'],
            sequence=stopTimes['stop_seq']
            )
    fp.write(str)
    
def convertExcelSheet(book, routeLongName, sheet_name, fp_r, fp_tr, fp_tm):
    # シートを取得
    sheet = book.sheet_by_name(sheet_name)

    routeShortName = ''
    
    # ルートID
    routeId = sheet_name.encode('utf-8')
    
    # セル：系統番号
    cell = sheet.cell(ROW_DATA_START, COL_KEITOU)    
    if cell.ctype == xlrd.XL_CELL_TEXT:
        routeShortName = cell.value.encode('utf-8')

    #　routes.txt
    routeType = 3
    outputRoute(fp_r, routeType, routeId, routeShortName, routeLongName)

    # 列：停留所名
    busStopList = []
    for cell in sheet.col(COL_BUSSTOPS)[ROW_DATA_START:]:
        if cell.ctype == xlrd.XL_CELL_TEXT:
            str = cell.value.encode('utf-8')
            stopId = findBusStop(str)
            if stopId is not None:
                busStopList.append(stopId)
            else:
                print("ERROR BusStopID")
    
    ## trips.txt, stop_times.txt
    trip = {}
    trip['shortName'] = routeShortName
    trip['route_id'] = routeId
    trip['headsign'] = routeLongName
    trip['direction'] = 0
    
    tripCount = 0
    
    # 列：運行時刻
    for ncol in range(COL_DATA_START, sheet.ncols):
        cells = sheet.col(ncol) 
        if cells[ROW_HEADER].ctype == xlrd.XL_CELL_EMPTY:
            break
        elif cells[ROW_HEADER].ctype == xlrd.XL_CELL_NUMBER:
            tripCount = tripCount +1
            tripId = "{0}-{1:03d}".format(routeId, tripCount)
            trip['trip_id'] = tripId
        
            # 運行曜日判定
            if cells[ROW_DAYTYPE].ctype == xlrd.XL_CELL_TEXT:
                if cells[ROW_DAYTYPE].value == u'平日':
                    trip['service_id'] = SERVICE_WORKDAY
                elif cells[ROW_DAYTYPE].value == u'土日祝':
                    trip['service_id'] = SERVICE_HOLIDAY
            # 運行時刻
            stopTimes = {'trip_id':tripId}
            timeList = cells[ROW_DATA_START:]
            for i in range(len(timeList)):
                cell = timeList[i]
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    # HH:MM:SS
                    timeStr = cell.value
                    pos = timeStr.index(':')
                    hh = int(timeStr[:pos])
                    mm = int(timeStr[pos+1:])
                    stopTimes['arrival_time'] = "{0:02d}:{1:02d}:00".format(hh,mm)
                    stopTimes['departure_time'] = stopTimes['arrival_time']
                    stopTimes['stop_id'] = busStopList[i]
                    stopTimes['stop_seq'] = i
                else:
                    break
                outputStopTimes(fp_tm, stopTimes)
                
        outputTrip(fp_tr, trip)
        
if __name__ == "__main__":
    
    ## 停留所データ読み込み
    (path, file) = os.path.split(sys.argv[0])
    
    try:
        fp = open( os.path.join(path, 'stops.txt'), 'r' )
        reader = csv.reader(fp)
        for row in reader:
            bsid = row[0]
            bsname = row[2]
            busStops[bsname] = bsid
    except IOError:
        print("IOError: stops.txt read failed.")
    else:
        fp.close()
    
    ##
    files = []
    
    argc = len(sys.argv)
    if argc == 1:
        print("ERROR: Specify excel files or directory in parameters.")
        exit
    elif argc >= 2:
        for arg in sys.argv[1:]:
            abspath = os.path.abspath(arg)
            if os.path.isdir(abspath):
                files = os.listdir(abspath)
            elif os.path.isfile(abspath):
                files.append(abspath)

    for filepath in files:
        if os.path.isfile(filepath):
            filename = os.path.basename(filepath)
            if filename[-4:] == '.xls':
                convertExcelFile(filepath)