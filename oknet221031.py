# coding:utf-8
from multiprocessing.resource_sharer import stop
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
import os
import signal
import gspread
from google.oauth2.service_account import Credentials
import yaml

with open('config.yml', 'r') as yml:
    config = yaml.safe_load(yml)
    myid = config['myid']
    mypass = config['mypass']
    SPREADSHEET_KEY  = config['SPREADSHEET_KEY']
    CHROMEDRIVER = config['CHROMEDRIVER']

def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


#スプレッドシート認証
scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
credentials = Credentials.from_service_account_file("credentials.json", scopes=scope)
gc = gspread.authorize(credentials)

# スプレッドシート（ブック）を開いて前準備
ws = gc.open_by_key(SPREADSHEET_KEY).sheet1
product_arry = ws.col_values(5)
value_arry = ws.col_values(2)
product_dict=(dict(zip(product_arry,value_arry)))

# ブラウザを開く。
chrome_service = fs.Service(executable_path=CHROMEDRIVER)
driver = webdriver.Chrome(service=chrome_service)
wait = WebDriverWait(driver=driver, timeout=30)

# OKネットスーパーのlogin画面を開く
driver.get("https://ok-netsuper.com/login")
wait.until(EC.presence_of_all_elements_located)

# ログイン操作
mail = driver.find_element(By.ID, "username")
password = driver.find_element(By.ID, "password")
mail.send_keys(myid)
password.send_keys(mypass)
sleep(20)
# reCAPTHA認証を人が操作する時間待機


#　配送日時のページを開く
driver.get("https://ok-netsuper.com/delivery-timetable/delivery-timetable-select")
wait.until(EC.presence_of_all_elements_located)
# 配送日時選択
driver.find_element(By.XPATH,"/html/body/div[2]/div[1]/div[2]/div/form/table/tbody/tr[3]/td[2]").click()
driver.find_element(By.ID,"continue-shopping").click()
sleep(1)

# タブを開く
driver.execute_script("window.open();")
# 開いたタブをアクティブにする
driver.switch_to.window(driver.window_handles[1])
# okホーム開く
driver.get("https://ok-netsuper.com/")
wait.until(EC.presence_of_all_elements_located)


try:
    for  product_name in product_arry:
        value = product_dict[product_name]
        if is_int(value):
            # 商品を検索欄選択
            search_product = driver.find_element(By.CSS_SELECTOR, "input.header-bottom-search-input")
            # 検索欄をクリア
            search_product.clear()
            # 検索欄に商品名を記入
            search_product.send_keys(product_name)
            # ENTERを押す
            search_product.send_keys(Keys.ENTER)
            wait.until(EC.presence_of_all_elements_located)
            # print(product_name)
            # print(value)
            try:
                for j in range(int(value)):
                    # 検索結果の一番目を個数分カートに入れる
                    driver.find_elements(By.CSS_SELECTOR,".ok-button")[4].click()
                    sleep(1)
            except IndexError:
                    # none_link = none_link + " " + product_name
                    print(product_name)
finally:
    os.kill(driver.service.process.pid,signal.SIGTERM)

driver.close
driver.quit