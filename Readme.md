複数のエクセルファイル、またはCSVファイルを1つのエクセルファイルに結合するスクリプト。
元のファイル名がそのまま結合先エクセルファイルの各シート名になる。

## 準備（初回のみ）
pip3 install openpyxl

## 実行
python3 marge.py

## 注意
* 作業前は output フォルダ内を空にする
* 作業後は input フォルダ内を空にする（前回実行時の input が output.xlsx に反映されてしまうため）
* input は xlsx ファイルか csv ファイルのみ入れる（xlsは不可）
