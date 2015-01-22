#!/usr/local/bin/python
# -*- coding: utf-8 -*-

'''
道南バスの停留所データのExcelファイルを読み込んでJSON形式に変換して出力する。

出力データフォーマット：
    1.停留所データ
        ID : String; 番号
        POLE : String; 柱の記号 'A', 'B'
        NAME : String; 停留所名称
        KEITOU : Array[ Int ]; 通過する系統
        POSITION : { "LON" : Float, "LAT" : Float }
'''

import sys
import os
import xlrd
import json

# Excelファイルのレイアウト
ROW_TOP = 0
ROW_DATA = 1

COL_NAME = 1
COL_ID = 8
COL_POLE = 12

# JSON出力のキー
_NAME = 'name'
_ID = 'id'
_POLE = 'pole'

def convertExcelFile(filepath):
    (path, filename) = os.path.split(filepath)
    
    busstopList = []
    
    ## Excelデータ読み取り
    book = xlrd.open_workbook(filepath, formatting_info=False)    
    for sheet in book.sheet_names():
        data = convertExcelSheet(book, sheet)
        busstopList.append(data)

    ## JSON出力
    newfilename = "BusStopList.json"
    newfile = os.path.join(path, newfilename)
    fp = open(newfile,'w')
    json.dump(busstopList, fp, ensure_ascii=False, indent=2)
    fp.close()

def convertExcelSheet(book, sheet_name):
    # シートを取得
    sheet = book.sheet_by_name(sheet_name)
    
    ## 系統ー停留所データ作成
    busstop = {}
    
    busstop[_NAME] = sheet.cell(ROW_DATA, COL_NAME).value.encode('utf-8')
    busstop[_ID] = int(sheet.cell(ROW_DATA, COL_ID).value)
    busstop[_POLE] = sheet.cell(ROW_DATA, COL_POLE).value.encode('utf-8')
    
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
            if ext == '.xls' or ext == '.xlsm':
                convertExcelFile(filepath)