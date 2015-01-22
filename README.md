# MuroranOpenData
むろらんオープンデータライブラリ <http://www.city.muroran.lg.jp/main/org2260/odlib.php> で公開されている
道南バス時刻表のExcelファイルをJSONファイルに変換する Pythonスクリプト。

Excelマクロで運行曜日の行を追加したレイアウトに対応している。公開されているオリジナルのレイアウトには対応いていない。

## Pythonスクリプト
スクリプトの引数には.xlsファイルまたはファイルが格納されているディレクトリのパスを指定する。

呼び出し例：
'python convert_excel_bus_data.py 室蘭市内線通過時刻表/*.xls'

##実行環境
Python 2.7系で動作確認。
外部モジュールとして xlrd <http://www.python-excel.org> を importしています。

## TODO

1. BusStop_*.jsonに含まれる停留所名称のリストを ID, Pole に置き換える。
2. BusStopList.jsonの停留所データに経緯度座標，通過系統のリストを追加する。

## Excelマクロ

公開されているExcelファイルでは運行曜日が「セルの背景色」で表現されている。
平日はシアン，土日祝はマゼンタ。

Pythonから背景色を判別出来ないため，マクロを用いて元のExcelファイルに運行曜日を「セルの値」として追加する。
