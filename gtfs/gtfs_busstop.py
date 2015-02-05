#!/usr/local/bin/python
# -*- coding: utf-8 -*-

'''
道南バスの停留所データのExcelファイルを読み込んでstops.txtを出力する。

出力フィールド：
        stop_id
        stop_code
        stop_name
        stop_desc
        stop_lat
        stop_lon
        stop_url
        location_type
        parent_station
        wheelchair_boarding
'''

import sys
import os
import xlrd
import csv
import json

output_fields = "stop_id,stop_code,stop_name,stop_desc,stop_lat,stop_lon,location_type,parent_station,wheelchair_boarding,stop_url"

# Excelファイルのレイアウト
ROW_TOP = 0
ROW_DATA = 1

COL_NAME = 1
COL_ID = 8
COL_POLE = 12

_NAME = "bsname"
_POLE = "bspole"
_ID = "bsid"

busStopData = None

def loadBusStops(filepath):
    '''
    # 国交省　国土数値データ　を GeoJsonで出力したファイルを読み込む
    '''
    
    fp = open(filepath, 'r')
    try:
        jsonData = json.load(fp)
    except :
        print("Error load geojson.", sys.exc_info()[0])
        jsonData = None
    else:
        fp.close()
    
    if jsonData is not None:
        # 停留所名称をキーにして座標の辞書を生成
        data = {}
        for x in jsonData['features']:
            key = x['properties']['P11_001']
            pos = x['geometry']['coordinates'][0]
            data[key] = pos
        
    return data
    
def getBusstopPosition(stopname):
    '''
    国土数値データと名称を付き合わせる
    ・ポール毎ではなく代表点の座標
    '''
    
    pos = {'lat':42.348363, 'lon':141.025856} # 東室蘭駅を初期値にしておく
    if busStopData is None:
        return pos
    
    if busStopData.has_key(stopname):
        pos['lat'] = busStopData[stopname][0]
        pos['lon'] = busStopData[stopname][1]
    else:
        # 国土数値データの名称と一致しない場合の処理
        
        # TODO: マルチバイト（１丁目）→シングルバイト（1丁目）
        
        # 省略への対応
        newkey = stopname
        if stopname[-3:] == u'学校前':
            newkey = stopname[:-3] + u'前'
            if busStopData.has_key(newkey):
                pos['lat'] = busStopData[newkey][0]
                pos['lon'] = busStopData[newkey][1]
        elif stopname[-1:] == u'前':
            newkey = stopname[:-1]
            if busStopData.has_key(newkey):
                pos['lat'] = busStopData[newkey][0]
                pos['lon'] = busStopData[newkey][1]
        elif stopname[-2:] == u'丁目':
            newkey = stopname[:-2]
            if busStopData.has_key(newkey):
                pos['lat'] = busStopData[newkey][0]
                pos['lon'] = busStopData[newkey][1]

    return pos


def convertExcelFile(filepath):
    '''
    stops.txt ファイルに出力する。
    '''
    
    (path, filename) = os.path.split(filepath)

    newfilename = "stops.txt"
    newfile = os.path.join(path, newfilename)
    fp = open(newfile,'w')
    fp.write(output_fields) # ヘッダー行出力
    fp.write('\n')
    
    unknownCount = 0
    
    # Excelデータ読み取り
    book = xlrd.open_workbook(filepath, formatting_info=False)    
    for sheet in book.sheet_names():
        
        # シートから停留所データを取り出し
        data = convertExcelSheet(book, sheet)
        
        # 停留所名称から座標を取得
        pos = getBusstopPosition(data[_NAME])
        if pos['lat'] == 0:
            unknownCount = unknownCount +1
        
        # 停留所を 1行ずつ出力
        locationType = 0
        wheelchair = 0
        parentStation = 0
        stopCode = ''
        stopUrl = '""'
        stopDesc = '""'
        
        str = "{id}_{pole},{code},{name},{desc},{lat},{lon},{type},{parent},{wheelchair},{url}\n".format(
            id=data[_ID],
            pole=data[_POLE].encode('utf-8'),
            name=data[_NAME].encode('utf-8'),
            code=stopCode,
            desc=stopDesc,
            lat=pos['lat'], 
            lon=pos['lon'],
            type=locationType,
            parent=parentStation,
            wheelchair=wheelchair,
            url=stopUrl)
        fp.write(str)
        
    fp.close()
    
    print( unknownCount, book.nsheets)


def convertExcelSheet(book, sheet_name):
    '''
    １シートが停留所のポールに対応する。
    '''
    
    # シートを取得
    sheet = book.sheet_by_name(sheet_name)
    
    busstop = {}
    
    busstop[_NAME] = sheet.cell(ROW_DATA, COL_NAME).value
    busstop[_ID] = int(sheet.cell(ROW_DATA, COL_ID).value)
    busstop[_POLE] = sheet.cell(ROW_DATA, COL_POLE).value
    
    return busstop
    
if __name__ == "__main__":
    
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
            ext = filename[ filename.rfind('.') : ]
            
            if ext == '.geojson':
                # 国土数値データを読み込む
                busStopData = loadBusStops(filepath)
            elif ext == '.xls' or ext == '.xlsm':
                # 停留所時刻表データを処理する
                convertExcelFile(filepath)