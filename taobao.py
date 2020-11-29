from selenium import webdriver as wb
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait as wdw
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pyautogui
import pyperclip
import random
import re
import getpass


# 大綱：1.因淘寶針對selenium開啟會有登入檢測，若使用selenium下開啟淘寶，手動驗證不會通過，因此調用pyautogui進行鍵盤和滑鼠自動化操作登入，接著爬蟲。
#       2.網路上資料提到有機器人檢測，若系統判定非人為，可以採用圖形驗證，但目前因為未出現此種情況，因此先使用sleep()來做間歇性爬取。
#       3.使用re分析網頁原始碼
def main():
    #user_name = str(input("請輸入淘寶帳號："))
    #pass_word = getpass.getpass("輸入淘寶密碼：")
    key_word = str(input("請輸入需要搜尋的項目："))
    page = int(input("請輸入搜尋頁數：\n（請在enter前將輸入法切換成英文）"))
    browser = login_taobao("", "", key_word)  # 登入
    #暫停兩分鐘進行手機驗證(use while needed)
    #time.sleep(120)
    browser = OpenPage(browser, "手錶")  # 搜尋動作
    # 機器人驗證
    # check_robot()
    top_last = []
    other_last = []
    for i in range(page):
        print("第{}頁".format(i + 1))
        time.sleep(3)  # 暫停防止被鎖
        page_str = browser.page_source
        result_top, result_other = taobao_spider(page_str)  # 爬蟲：
        top_last.extend(result_top)
        other_last.extend(result_other)
        browser = NextPage(browser)  # 換頁
        print(["name", "price", "sell"])  # 結果
        for i in range(len(result_top)):
            print(result_top[i])
        print(["name", "price", "sell", "shop", "location"])
        for i in range(len(result_other)):
            print(result_other[i])
    csv_file(top_last, key_word + "top", page,
             ["name", "price", "sell", "page"])
    csv_file(other_last, key_word + "other", page,
             ["name", "price", "sell", "shop", "location", "page"])

    browser.quit()


# 登入淘寶程序
def login_taobao(username, password, key_word):
    options = wb.ChromeOptions()
    options.add_experimental_option('excludeSwitches',
                                    ['enable-automation'])  # 切換到開發者模式
    browser = wb.Chrome(options=options)
    browser.get('https://login.taobao.com/member/login.jhtml')
    # 使用pyautogui自動登入淘寶（破解淘寶防selenium 機制）

    pyautogui.PAUSE = 0.5  # 設定每個動作間停頓，以防網頁速度跟不上
    pyautogui.FAILSAFE = True  # 啟用自動防故障功能
    width, height = pyautogui.size()  # 螢幕的寬度和高度
    str_user = pyperclip.copy(username)
    pyautogui.hotkey("ctrl", "command", "f")
    pyautogui.moveTo(1016, 375)
    pyautogui.click()
    #pyautogui.typewrite(pyperclip.paste(str_user), interval=0.25)
    pyautogui.hotkey("command", "v")
    pyautogui.moveTo(1011, 441)
    pyautogui.click()
    pyautogui.typewrite(password, interval=0.25)
    pyautogui.moveTo(953, 513)
    pyautogui.mouseDown()
    try:
        pyautogui.moveTo(1253, 513)
        pyautogui.mouseUp()
        left, top, width, height = pyautogui.locateOnScreen(
            "/Users/chang-yengtasi/Desktop/press5.PNG")
        print('綠色拉條驗證')
        pyautogui.moveTo(976, 576)
        pyautogui.click()
    except:
        print("無驗證")
        pyautogui.moveTo(953, 513)
        pyautogui.click()
    return browser


# 下一頁
def NextPage(browser):
    wait = wdw(browser, 10)
    wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'li[class="item next"]'))).click()
    return browser


# 打開要搜尋的頁面
def OpenPage(browser, keyword):
    #進行搜尋:
    wait = wdw(browser, 10)
    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'a[href="//www.taobao.com/"]'))).click()
    wait = wdw(browser, 10)
    input_key = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#mq')))
    click = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"]')))
    input_key.send_keys(keyword)
    click.click()
    # wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'iframe[class="srp-iframe"]'))).click() #打開新視窗
    #switchWindows(browser,1)  #切换視窗

    return browser


# 爬蟲
def taobao_spider(page_str):
    # 取元素
    other_sell_pattern = "<div class=\"item J_MouserOnverReq[\s\S]*?</li>"  # 其他所有
    top_sell_pattern = "<li class=\"[\s\S]*?item oneline[\s\S]*?</li>"  # 熱銷榜
    other_sell_list = re.findall(other_sell_pattern, page_str)
    top_sell_list = re.findall(top_sell_pattern, page_str)
    other_sell_result = other_sell(other_sell_list)
    top_sell_result = top_sell(top_sell_list)
    return top_sell_result, other_sell_result


# 推薦商品
def top_sell(top_sell_list):
    temp2_list = []
    for item in top_sell_list:
        name = item.split("title=\"")[1].split("\">")[0]  # put name
        price = float(item.split("</em>")[1].split("</a>")[0])  # put price
        sell = int(
            item.split("销量: <em>")[1].split("</em>")[0])  # put selling volume
        temp2_list.append([name, price, sell])
    return temp2_list


# 其他商品
def other_sell(other_sell_list):
    temp_list = []
    for item in other_sell_list:
        result = item.split("<div class=\"row row-2 title\"")[1].split(
            "</div>")[0]
        #有<span class="baoyou-intitle icon-service-free"></span>出現的情況
        try:
            result = result.split(
                "<span class=\"baoyou-intitle icon-service-free\"></span>")[1]
        except:
            result = result.split("<a id=\"[\s\S]*?\">")[0]
        try:
            result = result.split("\"\">")[1].strip()
        except:
            result = result.strip()
        result = result.replace('<span class="H">',
                                "").replace('</span>',
                                            "").replace("</a>",
                                                        "").strip()  # 加入名稱
        price = float(
            item.split("<div class=\"row row-1 g-clearfix\">")[1].split(
                "<strong>")[1].split("</strong>")[0])  # 加入價錢
        people = item.split("<div class=\"row row-1 g-clearfix\">")[1].split(
            "<div class=\"deal-cnt\">")[1].split("人付款")[0]  # 加入付款人數
        shop = item.split("<div class=\"row row-3 g-clearfix\">")[1].split(
            "<span>")[1].split("</span>")[0]  # 加入shop
        location = item.split("<div class=\"row row-3 g-clearfix\">")[1].split(
            "<div class=\"location\">")[1].split("</div>")[0]  # 加入地點
        temp_list.append([result, price, people, shop, location])
    return temp_list


# 輸出成csv檔案
def csv_file(list_name, file_name, page, col_list):
    outfile = open("/Users/chang-yengtasi/intern/" + file_name + ".csv", "w")
    for col in col_list:  # column index
        print(col, end=",", file=outfile)
    print(file=outfile)
    for i in range(page):
        for row in list_name:
            for k in row:
                print(k, end=",", file=outfile)
            print(i + 1, end=",", file=outfile)
            print(file=outfile)
    outfile.close()


main()
