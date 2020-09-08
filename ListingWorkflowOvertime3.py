import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from bs4 import BeautifulSoup

# chromeからdesknet'sを開く
driver = webdriver.Chrome("./chromedriver")
driver.get("https://dkn.e-omc.jp/cgi-bin/dneo/dneo.cgi?")

# id
id = driver.find_element_by_name("UserID")
id.send_keys("*****")

# pw
pw = driver.find_element_by_name("_word")
pw.send_keys("*****")

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

# 明細行ループ
for i in range(len(wfTblThElem)):

   # 状況が"取消し"の場合、スキップ
   wfTblThStatusElem = wfTblThElem[i].find_element_by_class_name("flow-td-status") # 状況element取得
   wfStatus = wfTblThStatusElem.get_attribute("title")

   if wfStatus == "取消し":
      continue

   # チェックボックスのidを取得
   wfTblThChkElem = wfTblThElem[i].find_elements_by_class_name("co-chk")  # チェックボックスelement取得
   wfTblThChkInputElem = wfTblThChkElem[0].find_element_by_name("id")  # チェックボックス配下のinput取得
   wfId = wfTblThChkInputElem.get_attribute("value")

   # 取得したidで単票表示
   driver.get("https://dkn.e-omc.jp/cgi-bin/dneo/zflow.cgi?cmd=flowindex#cmd=flowdisp&id=" + wfId)
   time.sleep(1)

   # ワークフロー種別が時間外申請以外はスキップ
   if driver.find_element_by_class_name("jco-cab-title") != "時間外申請書":
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

   
   print(wfMakeAplycant)
   print(wfMakeAplydate)
   print(wfDate)
   print(wfName)
   print(wfStartTime)
   print(wfEndTime)
   print(wfSvChk)
   print(wfOvertime)
   print(wfOvetimeMidnight)






# driver.close()
# driver.quit()