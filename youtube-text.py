####
# 処理の流れ
# （１）ページURLを取得
# （２）URLを開く
# （３）文字起こしボタンを押下
# （４）タイムラインを非表示
# （５）ページ要素解析
# （６）文字起こしテキストを取得
# （７）テキストファイルを作成
# （８）テキストファイルに貼り付け
# （９）テキストをダウンロード
# （１０）完了メッセージ
# ###

import time
import openpyxl
from openpyxl.styles.fonts import Font
from selenium import webdriver
import chromedriver_binary 
from selenium.webdriver import chrome
from selenium.webdriver.chrome import options
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Seleniumをあらゆる環境で起動させるChromeオプション
options = Options()
#options.add_argument('--headless'); #headlessモードを使用する（ヘッドレスモードとは、画面やページ遷移を表示せずに動作するモードです
options.add_argument('--disable-gpu'); #headlessモードで暫定的に必要なフラグ(そのうち不要になる)
options.add_argument('--disable-extensions'); #全ての拡張機能を向こうにする。ユーザスクリプトも無効にする
options.add_argument('--proxy-server="direct://"'); # Proxy経由ではなく直接接続する
options.add_argument('--proxy-bypass-list=*'); #全てのホスト名
options.add_argument('--start-maximized'); #起動時にウインドウを最大化する

#WebDriverのパスを指定(絶対でも相対パスでも可)
#DRIVER_PATH = r'C:\Sample\django\News\chromedriver_win32\chromedriver.exe' #ローカルはクロームのバージョンと互換しなかった

try:

    #ブラウザの起動
    #driver = webdriver.Chrome(executable_path=DRIVER_PATH, chrome_options=options)
    driver = webdriver.Chrome(chrome_options=options)
    # 要素が見つかるまで、最大10秒間待機する
    driver.implicitly_wait(5)

    # webページにアクセスする ここを取得するようにしたい
    url = 'https://www.youtube.com/watch?v=VyzqHFdzBKg'
    driver.get(url)

    #動画のタイトルを取得
    #container > h1 > yt-formatted-string
    selecter0 = '#container > h1 > yt-formatted-string'
    element = driver.find_element_by_css_selector(selecter0)
    title = element.text
    print("動画タイトル" + title)

    #ページ操作 https://posipochi.com/2021/05/03/python-scraping-css-selector/
    selecter1 = 'button[aria-label^="その他の操作"]' #対象要素のCSSセレクタ その他の操作ボタンを指定
    element = driver.find_element_by_css_selector(selecter1) #要素を検索
    print("element出力")

    element.click() #クリック処理

    #文字起こしボタン
    selecter2 = '#items > ytd-menu-service-item-renderer > tp-yt-paper-item > yt-icon' 
    element = driver.find_element_by_css_selector(selecter2) #要素を検索
    print("文字起こしボタン")
    element.click() #クリック処理

    time.sleep(5)

    #文字起こしのテキストを取得していく
    #タイムラインは消さなくてもできるけど消しても早そうと思ったけど、厳密なボタン要素ではないのでclickで動かん
    # selecter3 = 'div#panels > ytd-engagement-panel-section-list-renderer > div#header > ytd-engagement-panel-title-header-renderer > div#header > div#menu > ytd-menu-renderer > yt-icon-button'
    # element = driver.find_element_by_css_selector(selecter3) #要素を検索
    # print("文字起こしウィジェットその他ボタンクリック")
    # element.click() #クリック処理
    # time.sleep(5)

    #エクセルを新規作成
    book = openpyxl.Workbook()
    sheet = book.worksheets[0]
    sheet['A1'] = title
    sheet['A1'].font = Font(size=36)

    #文字起こしのテキストを取得していく。複数ある
    count = 1    
    while(True):
        selecter3 = '#body > ytd-transcript-body-renderer > div:nth-child(' + str(count) + ') > div.cues.style-scope.ytd-transcript-body-renderer > div.cue'
        if(len(driver.find_elements_by_css_selector(selecter3)) >0):
            ele = driver.find_element_by_css_selector(selecter3)
            print(str(count) + ":")
            print(ele.text)

            #エクセル書き込み
            sheet['A' + str(count+1)] = str(count)
            sheet['B' + str(count+1)] = ele.text

            count +=1

            if(count == 10):
                break
        else:
            break
    book.save(title + '.xlsx')

except TimeoutException as e:
    print(e)