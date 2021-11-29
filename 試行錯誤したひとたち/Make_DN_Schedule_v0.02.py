import pathlib  #標準ライブラリ
import os       #標準ライブラリ
import openpyxl #外部ライブラリ


"""
作業ディレクトリの確認
今はデスクトップ直下の「python￥wrk」で固定にしてる。
ある程度できたらinput使ってプロンプト上で入力するかファイルを読み込むか
"""
os.chdir('.\\python\wrk')
wrk_dir = os.getcwd()
print(wrk_dir +" 配下のファイルに対して処理を実行します。\n")

"""材料となるファイル名とパスを取得"""
path = pathlib.Path(wrk_dir)
for pass_obj in path.iterdir():
    pass_obj.match("店舗留置き*店着スケジュール.xls")
print("ファイル名を確認。")

wb = openpyxl.load_workbook(pass_obj)

"""ファイル名を取得してオープン"""
dname = os.path.dirname(pass_obj)
org_fname = os.path.basename(pass_obj)
org_fname1 = os.path.splitext(os.path.basename(pass_obj))[0]
print(org_fname+"　を加工します。")

wb.save(dname+ "\\" + org_fname1 +"（D+N日表記）"+".xlsx")
"""
シート操作のフローが悩ましい
空のシート作成=
    =>元シートの内容を全コピー(単純なシートコピーの仕方がわからん
        =>シートの名前を正しいものに
            =>必要に応じて新潟修正
            　関数をぶちこむ

ws = wb.get_sheet_by_name('SHEETNAME')
ws.title = 'ＳＥＪ店舗'

ws = wb.get_sheet_by_name('SHEETNAME')
ws.title = 'その他'

ws = wb.get_sheet_by_name('SHEETNAME')
ws.title = 'ＳＥＪ店舗(基表)'

ws = wb.get_sheet_by_name('SHEETNAME')
ws.title = 'その他（基表）'

wb.copy_worksheet(ws)
https://qiita.com/yu_yagishita/items/fd3abf6a53d042f1b782
ＳＥＪ店舗 (新潟修正)
その他 (新潟修正)
その他 (新潟修正)
ＳＥＪ店舗(基表) (新潟修正)
その他（基表） (新潟修正)
ＳＥＪ店舗 (差分)
その他 (差分)
ＳＥＪ店舗 (D+N)
その他 (D+N)
"""
#ＳＥＪ店舗 (差分)
#=IF(DATE(MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,1,4),MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,5,2),MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,7,2))<DATE(MID('ＳＥＪ店舗 (新潟修正)'!C6,1,4),MID('ＳＥＪ店舗 (新潟修正)'!C6,5,2),MID('ＳＥＪ店舗 (新潟修正)'!C6,7,2)),DATEDIF(DATE(MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,1,4),MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,5,2),MID('ＳＥＪ店舗(基表) (新潟修正)'!C6,7,2)),DATE(MID('ＳＥＪ店舗 (新潟修正)'!C6,1,4),MID('ＳＥＪ店舗 (新潟修正)'!C6,5,2),MID('ＳＥＪ店舗 (新潟修正)'!C6,7,2)),  "D"),)

#その他 (差分)
#=IF(DATE(MID('その他（基表） (新潟修正)'!C6,1,4),MID('その他（基表） (新潟修正)'!C6,5,2),MID('その他（基表） (新潟修正)'!C6,7,2))<DATE(MID('その他 (新潟修正)'!C6,1,4),MID('その他 (新潟修正)'!C6,5,2),MID('その他 (新潟修正)'!C6,7,2)),DATEDIF(DATE(MID('その他（基表） (新潟修正)'!C6,1,4),MID('その他（基表） (新潟修正)'!C6,5,2),MID('その他（基表） (新潟修正)'!C6,7,2)),DATE(MID('その他 (新潟修正)'!C6,1,4),MID('その他 (新潟修正)'!C6,5,2),MID('その他 (新潟修正)'!C6,7,2)),  "D"),)

#ＳＥＪ店舗 (D+N)
#=IF('ＳＥＪ店舗 (差分)'!C6=0,'ＳＥＪ店舗(基表) (新潟修正)'!C6,'ＳＥＪ店舗(基表) (新潟修正)'!C6&"+"&'ＳＥＪ店舗 (差分)'!C6)

#その他 (D+N)
#=IF('その他 (差分)'!C6=0,'その他（基表） (新潟修正)'!C6,'その他（基表） (新潟修正)'!C6&"+"&'その他 (差分)'!C6)
