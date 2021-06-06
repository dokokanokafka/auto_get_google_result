# !python3
import openpyxl
import time
from selenium import webdriver

#ここで調べたいキーワード、取得件数を設定
#googleで検索する文字(キーワードが複数の場合、スペースで区切る)
search_string = '京都　不動産'
#取得件数
search_number = 50

#Seleniumを使うための設定とgoogleの画面への遷移
INTERVAL = 2.5
URL = "https://www.google.com/"
# chromedriverが同じ階層にある場合は
# driver_path = "./chromedriver"
# driver = webdriver.Chrome(executable_path=driver_path)

driver = webdriver.Chrome(executable_path='/Users/haruhikido/Desktop/chromedriver')

time.sleep(INTERVAL)
driver.get(URL)
time.sleep(INTERVAL)

#文字を入力して検索
# Googleの画面を開いた後、name属性が’q’の要素(検索テキストボックス)に検索文字列を入力
driver.find_element_by_name('q').send_keys(search_string)
# ’btnK’(’btnK’は検索ボタン)をクリック
driver.find_elements_by_name('btnK')[1].click() #btnKが2つあるので、その内の後の方
time.sleep(INTERVAL)

#検索結果の一覧を取得する
# 検索結果をn件分resultsリストに格納
results = []
flag = False
# 件数がn件に到達するまで無限ループ
while True:
    # g_aryは検索結果のタイトルとURLが入った要素をリストで保持
    g_ary = driver.find_elements_by_class_name('g')
    # for inでループし、タイトルとURLを中から取り出してresult辞書に格納(PHPのforeachと同じ)後、resultsに一つずつ渡す
    for g in g_ary:
        result = {}
        result['url'] = g.find_element_by_class_name('yuRUbf').find_element_by_tag_name('a').get_attribute('href')
        result['title'] = g.find_element_by_tag_name('h3').text
        # result['title'] = g.find_element_by_tag_name('h3').find_element_by_tag_name('span').text
        results.append(result)
        # ループが終わってるか判定
        if len(results) >= search_number:
            flag = True
            break
    # 終わってればループを終了
    if flag:
        break
    # ループが終わってなければ次ページへ
    driver.find_element_by_id('pnnext').click()
    time.sleep(INTERVAL)

#ワークブックの作成とヘッダ入力
workbook = openpyxl.Workbook()
# sheetシートを取得
sheet = workbook.active
sheet['A1'].value = search_string
sheet['B1'].value = search_number
sheet['C1'].value = '件'
sheet['A3'].value = 'タイトル'
sheet['B3'].value = 'URL'

#シートにタイトルとURLの書き込み(4行目から)
for row, result in enumerate(results, 4):
    sheet[f"A{row}"] = result['title']
    sheet[f"B{row}"] = result['url']

# 同じ階層に検索結果のExcelファイルを保存
workbook.save(f"google_search_{search_string}.xlsx")
driver.close()