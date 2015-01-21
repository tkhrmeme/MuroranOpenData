# MuroranOpenData
むろらんオープンデータライブラリ <http://www.city.muroran.lg.jp/main/org2260/odlib.php> で公開されている
道南バス時刻表のExcelファイルをJSONファイルに変換する Pythonスクリプト。

## Pythonスクリプト
スクリプトの引数には.xlsファイルまたはファイルが格納されているディレクトリのパスを指定する。

呼び出し例：
'python convert_excel_bus_data.py 室蘭市内線通過時刻表/*.xls'

##実行環境
Python 2.7系で動作確認。
外部モジュールとして xlrd <http://www.python-excel.org> を importしています。

## TODO

1.Excelファイルから運行曜日が判別が出来ないためPythonスクリプトで出力される値は暫定となっている。
2.Excelマクロで拡張した値への対応は未定。

## Excelマクロ

公開されているExcelファイルでは運行曜日が「セルの背景色」で表現されている。
平日はシアン，土日祝はマゼンタ。

Pythonから背景色を判別出来ないため，マクロを用いて元のExcelファイルに運行曜日を「セルの値」として追加する。
