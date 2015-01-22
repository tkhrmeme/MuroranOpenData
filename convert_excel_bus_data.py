#!/usr/local/bin/python
# -*- coding: utf-8 -*-

'''
道南バスの時刻表データのExcelファイルを読み込んでJSON形式に変換して出力する。

出力データフォーマット：
    1.系統ー停留所データ
        KEITOU_NAME : String; ファイル名の()に挟まれた名称
        KEITOU_ID : String; シートの名称
        KEITOU_NUMBER : String シート２列目の系統番号。空文字列の場合が有る。
        KEITOU_BUSSTOP : Array[String]; シート１列目の駅名以下のリスト
    2.系統ー運行時刻データ
        KEITOU_SCHEDULE : Array[ Schedule ]; Scheduleオブジェクトのリスト
        Schedule : { 
                    SCHEDULE_DAYTYPE: Int; 運行曜日(0: 平日，1: 土日祝, 9:不明)
                    SCHEDULE_TIME: Array[ String ]; 時刻のリスト
                    }
    
'''

import sys
import os
import xlrd
import json

# Excelファイルのレイアウト
ROW_TOP = 0
ROW_HEADER = 1
ROW_DAYTYPE = 2
ROW_DATA_START = 3

COL_BUSSTOPS = 0
COL_KEITOU = 1
COL_DATA_START = 2

# JSON出力のキー
KEITOU_NAME = 'keitou_name'
KEITOU_ID = 'keitou_id'
KEITOU_NUMBER = 'keitou_number'
KEITOU_BUSSTOP = 'busstops'
KEITOU_SCHEDULE = 'schedule'

SCHEDULE_DAYTYPE = 'daytype'
SCHEDULE_TIME = 'time'

# 運行曜日
DAYTYPE_WORKDAY = 0 # 平日
DAYTYPE_HOLIDAY = 1 # 土日祝
DAYTYPE_UNKNOWN = 9 # 不明


def convertExcelFile(filepath):
    busdata = {}
    
    (path, filename) = os.path.split(filepath)
    filename = filename[:-4]
    routename = filename[filename.find('(')+1:filename.find(')')]

    busstopList = []
    scheduleList = []
    
    ## Excelデータ読み取り
    book = xlrd.open_workbook(filepath, formatting_info=True)    
    for sheet in book.sheet_names():
        keitou_data = convertExcelSheet(book, sheet, routename)
        busstopList.append(keitou_data['busstop'])
        scheduleList.append(keitou_data['schedule'])

    ## JSON出力
    newfilename = "BusStop_" +filename +'.json'
    newfile = os.path.join(path, newfilename)
    fp = open(newfile,'w')
    json.dump(busstopList, fp, ensure_ascii=False, indent=2)
    fp.close()

    newfilename = "Schedule_" +filename +'.json'
    newfile = os.path.join(path, newfilename)
    fp = open(newfile,'w')
    json.dump(scheduleList, fp, ensure_ascii=False, indent=2)
    fp.close()

def convertExcelSheet(book, sheet_name, route_name):
    # シートを取得
    sheet = book.sheet_by_name(sheet_name)
    
    data = {}
    
    # 行：ヘッダー
#     row_haeder = sheet.row(ROW_HEADER)
#     header = []
#     for cell in row_haeder[2:]:
#         if cell.ctype == xlrd.XL_CELL_TEXT:
#             header.append(cell.value)   
#         elif cell.ctype == xlrd.XL_CELL_NUMBER:
#             header.append(int(cell.value))
    
    ## 系統ー停留所データ作成
    keitou_busstop = {}
    keitou_busstop[KEITOU_NAME] = route_name
    keitou_busstop[KEITOU_ID] = sheet_name.encode('utf-8')
    
    # セル：系統番号
    cell = sheet.cell(ROW_DATA_START, COL_KEITOU)    
    if cell.ctype == xlrd.XL_CELL_TEXT:
        keitou_number = cell.value.encode('utf-8')
        keitou_busstop[KEITOU_NUMBER] = keitou_number

    # 列：停留所名
    keitou_busstop[KEITOU_BUSSTOP] = []    
    for cell in sheet.col(COL_BUSSTOPS)[ROW_DATA_START:]:
        if cell.ctype == xlrd.XL_CELL_TEXT:
            str = cell.value.encode('utf-8')
            keitou_busstop[KEITOU_BUSSTOP].append(str)
        
    data['busstop'] = keitou_busstop
    
    ## 系統ー運行時刻データ
    keitou_schedule = {}
    keitou_schedule[KEITOU_NAME] = route_name
    keitou_schedule[KEITOU_ID] = sheet_name.encode('utf-8')
    keitou_schedule[KEITOU_NUMBER] = keitou_number
    
    operations = []
    # 列：運行時刻
    for ncol in range(COL_DATA_START, sheet.ncols):
        cells = sheet.col(ncol) 
        if cells[ROW_HEADER].ctype == xlrd.XL_CELL_EMPTY:
            break
        elif cells[ROW_HEADER].ctype == xlrd.XL_CELL_NUMBER:
            op = {}
            # 運行曜日判定
            if cells[ROW_DAYTYPE].ctype == xlrd.XL_CELL_TEXT:
                if cells[ROW_DAYTYPE].value == u'平日':
                    op[SCHEDULE_DAYTYPE] = DAYTYPE_WORKDAY
                elif cells[ROW_DAYTYPE].value == u'土日祝':
                    op[SCHEDULE_DAYTYPE] = DAYTYPE_HOLIDAY
                else:
                    op[SCHEDULE_DAYTYPE] = DAYTYPE_UNKNOWN
            
            op[SCHEDULE_TIME] = []
            for cell in cells[ROW_DATA_START:]:
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    op[SCHEDULE_TIME].append(cell.value)
            operations.append(op)
            
    keitou_schedule[KEITOU_SCHEDULE] = operations
            
    data['schedule'] = keitou_schedule
    
    return data
    
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
            if filename[-4:] == '.xls':
                convertExcelFile(filepath)