from selenium import webdriver
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service as ChromeService
import csv
import time
import datetime
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
import json
import ssl
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart

mailaddress = "ここに送信先アドレスを記入"
now = time.time()
csvname = "takken-" + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S-%f") + ".csv"
#csv出力、リストで受け取る
def outputcsv(listdata):
    with open(csvname, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(listdata)

#ある業者ページから情報を抜く、記録のために都道府県コードももらう
def getinfo(driver, code):
    prefectureCode = code
    number = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/table/tbody/tr[1]/td").text
    companyName = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/table/tbody/tr[2]/td").text.split("\n")[1]
    director = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/table/tbody/tr[3]/td").text.split("\n")[1]
    address = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/table/tbody/tr[4]/td").text.replace("\n", "")
    phone = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/table/tbody/tr[5]/td").text
    money = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[1]/div/table/tbody/tr[2]/td").text
    permissionDate = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/div[2]/div[1]/div/table/tbody/tr[1]/td/a").text
    #業種
    normalindustrylist = list()
    specialindustrylist = list()
    for i in range(1, 30):
        industryValue = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/table[1]/tbody/tr[2]/td[" + str(i) + "]").text
        if industryValue == "1":
            normalindustrylist.append(driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[5]/div/table[1]/tbody/tr[1]/td[" + str(i) + "]").text.replace("\n", ""))
        elif industryValue == "2":
            specialindustrylist
    normalindustry = ",".join(normalindustrylist)
    specialindustry = ",".join(specialindustrylist)

    
    return [prefectureCode, number, companyName, director, address, phone, normalindustry, specialindustry, money, permissionDate]

#メール送信
def send_mail(mail_from, mail_to, mail_subject, mail_body, file_name, path):
    settings_file = open('mailsettings.json','r')
    settings_data = json.load(settings_file)
    host = settings_data["host"]
    port = settings_data["port"]
    account = settings_data["account"]
    password = settings_data["password"]

    """ メッセージのオブジェクト """
    msg = MIMEMultipart()
    msg['Subject'] = mail_subject
    msg['From'] = mail_from
    msg['To'] = mail_to
    msg.attach(MIMEText(mail_body, "plain", "utf-8"))
    
    with open(file_name, "r") as fp:
        attach_file = MIMEText(fp.read(), "plain")
        attach_file.add_header(
            "Content-Disposition",
            "attachment",
            filename = file_name
        )
        msg.attach(attach_file)
    # エラーキャッチ
    try:
        context = ssl.create_default_context()
        server = SMTP_SSL(host, port, context=context)
        server.login(account, password)
        server.send_message(msg)
        server.quit()


    except Exception:
        import traceback
        traceback.print_exc()
    
    return "メール送信完了"

#本体
def main():
    #今どこまでやったのか
    with open(csvname, mode='r') as file:
        reader = csv.reader(file)
        csvdata = list(reader)
    #最終行の都道府県コード、初期状態ならヘッダ
    lastPrefecture = csvdata[len(csvdata) - 1][0]
    #見つかったのがヘッダだった場合は別
    if lastPrefecture != "都道府県コード":
        #その都道府県をリセット
        #インデックスがずれないように、後ろから
        for i in range(len(csvdata) - 1, 1, -1):
            if csvdata[i][0] == lastPrefecture:
                csvdata.remove(csvdata[i])
        #書き換え
        with open(csvname, "w", newline="") as file:
            writer = csv.writer(file)
            writer.writerows(csvdata)
    else:
        lastPrefecture = 1
    
    #raspberrypi用
    """
    path = "/usr/bin/chromedriver"
    options = ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=ChromeService(path), options=options)
    """
    #windows用
    driver = webdriver.Chrome()
    
    driver.get("https://etsuran2.mlit.go.jp/TAKKEN/kensetuKensaku.do?outPutKbn=1")
    
    #コードは2桁の数字なので、0埋めが要る
    #1-47版
    prefectureCodes = [str(code).zfill(2) for code in range(int(lastPrefecture), 48) ]

    print(prefectureCodes)
    
    for i in prefectureCodes:
        #進捗出力
        print("now", i)
        #都道府県選択、ないものもある
        prefectureDropdown = Select(driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[4]/div/div[5]/div[2]/select[2]"))
        prefectureDropdown.select_by_value(i)

        #ソート
        sortDropdown = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[4]/div/div[6]/div[1]/select")
        Select(sortDropdown).select_by_value("2")
        #降順
        driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[4]/div/div[6]/div[1]/p/input[2]").click()
        #検索
        driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[4]/div/div[6]/div[5]/img").click()

        #ひとまず2ページ
        for j in range(1, 3):
            #ページ移動
            dropdown = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/div[6]/div[3]/select")
            Select(dropdown).select_by_value(str(j))

            #1つのページのレコード一覧
            table = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/table/tbody")
            records = table.find_elements(By.TAG_NAME, "tr")
            
            #レコード1つめはヘッダ
            k = 1
            for record in records[1:]:
                link = driver.find_element(By.XPATH, "/html/body/form/div/div[2]/table/tbody/tr[" + str(k + 1) + "]/td[4]/a")
                link.click()
                targetdata = getinfo(driver, i)
                outputcsv(targetdata)
                driver.back()
                k += 1


    driver.quit()
    send_mail("差し出し元", mailaddress, "test", "testdata", csvname, "")


#csvヘッダ
outputcsv(["都道府県コード", "許可番号", "会社名", "代表者名", "所在地", "電話番号", "業種(一般)", "業種(特殊)", "資本金", "許可年月日"])

#10回挑戦する
for i in range(10):
    try:
        main()
        exit()
    except Exception:
        import traceback
        traceback.print_exc()
    #エラーなら1分後再挑戦
    time.sleep(30)

#挑戦してダメだったとき
print("error")