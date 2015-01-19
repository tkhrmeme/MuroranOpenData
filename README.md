# MuroranOpenData
むろらんオープンデータライブラリ (http://www.city.muroran.lg.jp/main/org2260/odlib.php) で公開されている
道南バス時刻表のExcelファイルをJSONファイルに変換する Pythonスクリプト。

スクリプトの引数には.xlsファイルまたはファイルが格納されているディレクトリのパスを指定する。

呼び出し例：
python convert_excel_bus_data.py 室蘭市内線通過時刻表/*.xls


TODO:Excelファイルから運行曜日が判別が出来ないため出力される値は暫定となっている。
