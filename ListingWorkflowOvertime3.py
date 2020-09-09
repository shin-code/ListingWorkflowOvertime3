import time
import os
from selenium import webdriver
from selenium.webdriver.support.select import Select
from bs4 import BeautifulSoup
import openpyxl as px
from datetime import datetime
import myFn

# 定数定義
XLSX_FILE = "ListWorkflowOvertime3.xlsx"
SHEET_NAME = "Overtime3"

# パラメータ
dnId = input("ログインID >> ")
dnPw = input("パスワード >> ")
fromDate = myFn.text_to_date(input("抽出範囲日付FROM ( yyyy/mm/dd形式 ) >> "))  # 日付変換

'''
**************************************************
 出力先エクセルを開く
**************************************************
'''
# 出力先ファイル存在する場合はファイルを開く
if os.path.isfile(XLSX_FILE):

   wb = px.load_workbook(XLSX_FILE)
   ws = wb.get_sheet_by_name(SHEET_NAME)


else:  # 出力先ファイル存在しない場合はファイルを新規作成

   wb = px.Workbook()
   ws = wb.active
   ws.title = SHEET_NAME  # シート名

   # 作成情報
   ws['A1'].value = '時間外申請一覧（システム３課）'

   # タイトル行
   ws['A2'].value = '申請者'
   ws['B2'].value = '申請日時'
   ws['C2'].value = '日付'
   ws['D2'].value = '氏名'
   ws['E2'].value = '実績開始時間'
   ws['F2'].value = '実績終了時間'
   ws['G2'].value = '朝のサーバーチェック'
   ws['H2'].value = '申請時間'
   ws['I2'].value = '申請深夜時間'
   ws['J2'].value = '一覧作成日時'

   wb.save(XLSX_FILE)  # 一旦保存

'''
**************************************************
 desknet'sログイン → ワークフロー遷移
**************************************************
'''
# chromeからdesknet'sを開く
driver = webdriver.Chrome("./chromedriver")
driver.get("https://dkn.e-omc.jp/cgi-bin/dneo/dneo.cgi?")

# id
id = driver.find_element_by_name("UserID")
id.send_keys(dnId)

# pw
pw = driver.find_element_by_name("_word")
pw.send_keys(dnPw)

# ログイン
id.submit()

time.sleep(3)

# ワークフロー
driver.get("https://dkn.e-omc.jp/cgi-bin/dneo/zflow.cgi?cmd=flowindex")

time.sleep(3)

# ワークフローステータス -> 完了
wfStatusElem = driver.find_element_by_id("flow-list-type-sel")  # element取得
wfStatusObj = Select(wfStatusElem)
wfStatusObj.select_by_visible_text("完了")

time.sleep(1)

wfTblThElem = driver.find_elements_by_class_name("flow-list-line")  # 明細行elements取得

'''
**************************************************
 ワークフロー１件ずつ処理
**************************************************
'''
wsRow = ws.max_row # エクセル最終行取得

# 明細行ループ
for i in range(len(wfTblThElem)):

   wfTblTdElem = wfTblThElem[i].find_elements_by_tag_name("td")

   # 状況が"取消し"の場合、スキップ
   if wfTblTdElem[3].text == "取消し":
      continue

   # 作成日付を取得
   slashCount = wfTblTdElem[6].text.count("/")  # スラッシュの出現回数
   if slashCount == 1:  # 年が省略されているので今年を付加
      textDate = str(datetime.now().year) + "/" + wfTblTdElem[6].text[0:5]
   else:
      textDate =wfTblTdElem[6].text[0:10]
   wfMakeAplydate = myFn.text_to_date(textDate)  # 日付変換

   # 作成日付が抽出範囲日付FROM以下の場合、ループを抜ける
   if wfMakeAplydate < fromDate:
      break

   # チェックボックスのidを取得
   wfTblThChkElem = wfTblThElem[i].find_elements_by_class_name("co-chk")  # チェックボックスelement取得
   wfTblThChkInputElem = wfTblThChkElem[0].find_element_by_name("id")  # チェックボックス配下のinput取得
   wfId = wfTblThChkInputElem.get_attribute("value")

   # 取得したidで単票表示
   driver.get("https://dkn.e-omc.jp/cgi-bin/dneo/zflow.cgi?cmd=flowindex#cmd=flowdisp&id=" + wfId)
   time.sleep(1)

   # ワークフロー種別が時間外申請以外はスキップ
   if driver.find_element_by_class_name("jco-cab-title").text != "時間外申請書":
      continue

   """
   作成情報

   table (flow-view-meta)
        td(flow-view-meta-title)   td(flow-view-meta-data)
      +--------------------------+-----------------------------+
   tr | 申請組織                  | システム○課                   |
      +--------------------------+-----------------------------+
   tr | 申請者                    | 宮崎  一郎                   |
      +--------------------------+-----------------------------+
   tr | 申請日時                  | 2020年09月01日(木) 17:00      |
      +--------------------------+-----------------------------+
   tr | 決済状況                  | 完了                         |
      +--------------------------+-----------------------------+
   """
   wfAplyTblTdElem = driver.find_elements_by_class_name("flow-view-meta-title")  # 1列目element取得
   
   # 申請者と申請日時が何行目か調べる
   for i in range(len(wfAplyTblTdElem)):
      if wfAplyTblTdElem[i].text == "申請者":
         wfMakeAplycantRow = i
      elif wfAplyTblTdElem[i].text == "申請日時":
         wfMakeAplydateRow = i
    
   wfAplyTblTdElem = driver.find_elements_by_class_name("flow-view-meta-data")  # 2列目element取得

   # 申請者
   wfMakeAplycant = wfAplyTblTdElem[wfMakeAplycantRow].text
   # 申請日時
   wfMakeAplydate = wfAplyTblTdElem[wfMakeAplydateRow].text

   """
   明細部分の基本的な構成

   氏名        宮崎 一郎    ←1つのパーツ form-parts
    ↑          ↑
   fontタグ    fontタグ
   """
   wfFormParts = driver.find_elements_by_class_name("form-parts")

   # 日付
   wfFpFonts = wfFormParts[2].find_elements_by_tag_name("font")
   wfDate = wfFpFonts[1].text

   # 氏名
   wfFpFonts = wfFormParts[3].find_elements_by_tag_name("font")
   wfName = wfFpFonts[1].text

   # 実績開始時間
   wfFpFonts = wfFormParts[15].find_elements_by_tag_name("font")
   wfStartTime = wfFpFonts[0].text

   # 実績終了時刻
   wfFpFonts = wfFormParts[16].find_elements_by_tag_name("font")
   wfEndTime = wfFpFonts[1].text

   # 朝のサーバーチェック
   wfFpImg = wfFormParts[17].find_element_by_tag_name("img")
   if "form_checkbox_on.gif" in wfFpImg.get_attribute("src"):
      wfSvChk = True
   else:
      wfSvChk = False

   # 申告時間
   wfFpFonts = wfFormParts[18].find_elements_by_tag_name("font")
   wfOvertime = wfFpFonts[1].text

   # 深刻深夜時間
   wfFpFonts = wfFormParts[19].find_elements_by_tag_name("font")
   wfOvetimeMidnight = wfFpFonts[1].text

   # エクセル出力   
   ws.cell(row=wsRow, column=1).value = wfMakeAplycant      # 申請者
   ws.cell(row=wsRow, column=2).value = wfMakeAplydate      # 申請日時
   ws.cell(row=wsRow, column=3).value = wfDate              # 日付
   ws.cell(row=wsRow, column=4).value = wfName              # 氏名
   ws.cell(row=wsRow, column=5).value = wfStartTime         # 実績開始時間
   ws.cell(row=wsRow, column=6).value = wfEndTime           # 実績終了時間
   ws.cell(row=wsRow, column=7).value = wfSvChk             # 朝のサーバーチェック
   ws.cell(row=wsRow, column=8).value = wfOvertime          # 申請時間
   ws.cell(row=wsRow, column=9).value = wfOvetimeMidnight   # 申請深夜時間

   wsRow += 1  # エクセル書き込み行カウントアップ

wb.save(XLSX_FILE)  # エクセル保存

# driver.close()
# driver.quit()