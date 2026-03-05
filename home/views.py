from django.shortcuts import render, redirect
from datetime import datetime
from home.models import Post
from django.http import JsonResponse
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from PIL import Image
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

import time
import pytesseract
import re
import os

chromeWeb = None
game = None
environment = None
uploaded_file = None


# Create your views here.
def homepage(request):
    posts = Post.objects.all()
    now = datetime.now()
    return render(request, "home.html", locals())


def showpost(request, slug):
    try:
        post = Post.objects.get(slug=slug)
        if post != None:
            return render(request, "post.html", locals())
    except:
        return redirect("/")


def ppsg(requset):

    account = "帳號："
    pswd = "密碼："
    game = "遊戲類別："
    environment = "測試環境："
    excel = "上傳檔案："
    website = "網站："

    context = {
        "account": account,
        "pswd": pswd,
        "game": game,
        "environment": environment,
        "excel": excel,
        "website": website,
    }

    return render(requset, "ppsg.html", context)


def upload(request):
    global game, chromeWeb, environment, uploaded_file

    if request.method == "POST":
        account = request.POST.get("account")
        pswd = request.POST.get("pswd")
        game = request.POST.get("game")
        environment = request.POST.get("environment")
        website = request.POST.get("website")
        uploaded_file = request.FILES.getlist("excelFile")  # 獲取上傳的檔案
        
        parameters(website, environment, account, pswd)
        
        # 將資訊存入 session
        request.session['game'] = game
        request.session['environment'] = environment
        request.session['chrome_initialized'] = True
        
        
        # 檢查 chromeWeb 是否成功開啟
        if chromeWeb is None:
            return JsonResponse({"error": "無法開啟瀏覽器"}, status=500)
        # 等待登入
        time.sleep(3)
        return JsonResponse({"message": "檔案上傳成功！"}, status=200)

    else:
        return JsonResponse({"error": "不支持的請求方法"}, status=400)


# 參數判斷和URL設定
def parameters(website, environment, account, pswd):

    url_mapping = {
        (
            "thor_admin"
        ): "https://admin.12vin.com/(S(h2uux4srsyv0fwgu3ym2pmti))/default.aspx",
        (
            "thor_agent"
        ): "https://agent.12vin.com/(S(1clf4jfquob2frkapsd1egc3))/default.aspx",
        (
            "thor_max222agent"
        ): "https://max222agent.12vin.com/(S(kqtj2qkkbil4n0wgn5cy0npz))/default.aspx",
        ("thor_gcadmin"): "https://gameadmin.12vin.com/",
        (
            "sta1_admin"
        ): "https://admin.vina368.net/(S(eqjgza4zuwutd5y53y50m2w1))/default.aspx",
        (
            "sta1_agent"
        ): "https://cmdbetagent.368aa.net/(S(s0oh1niwwjpjtpdfnu0l0r5b))/default.aspx",
        (
            "sta1_max222agent"
        ): "https://max222agent.vina368.net/(S(ltsyfwcwxej34le5rrkoh2vf))/default.aspx",
        ("sta1_gcadmin"): "http://gcadmin.cmdbetsta.com/",
        (
            "sta2_admin"
        ): "https://admin.cmmd368.com/(S(szdtumkbydibusat1o2xqqd1))/default.aspx",
        (
            "sta2_agent"
        ): "https://cmdbetagent.cmmd368.com/(S(2ml21gvm0f0nwr1dkdqewr1j))/default.aspx",
        (
            "sta2_max222agent"
        ): "https://max222agent.cmmd368.com/(S(s5qklauyhfsbuhe5qxlxcd02))/default.aspx",
        ("sta2_gcadmin"): "https://gcadmin.cmmd368.com/",
        (
            "prod_admin"
        ): "https://admin.cmdbet.biz/(S(psil03jsmbweb24syq3pbbxi))/default.aspx",
        (
            "prod_agent"
        ): "https://agent.cmdbet.com/(S(xk3u4trawghnzzifbcyd3pib))/default.aspx",
        (
            "prod_max222agent"
        ): "https://agent.max222.com/(S(xdocse1gos51nz2rue5ao1ih))/default.aspx",
        ("prod_gcadmin"): "https://gcadmin.cmdbet.biz/",
    }

    url_key = f"{environment}_{website}"
    url = url_mapping.get(url_key)

    handlers = {
        "admin": {
            "thor": lambda: handle_thor(account, pswd, url),
            "sta1": lambda: handle_sta1(account, pswd, url),
            "sta2": lambda: handle_sta2(account, pswd, url),
            "prod": lambda: handle_prod(account, pswd, url),
        },
        "agent": {
            "thor": lambda: handle_thor(account, pswd, url),
            "sta1": lambda: handle_sta1(account, pswd, url),
            "sta2": lambda: handle_sta2(account, pswd, url),
            "prod": lambda: handle_prod(account, pswd, url),
        },
        "max222agent": {
            "thor": lambda: handle_thor(account, pswd, url),
            "sta1": lambda: handle_sta1(account, pswd, url),
            "sta2": lambda: handle_sta2(account, pswd, url),
            "prod": lambda: handle_prod(account, pswd, url),
        },
        "gcadmin": {
            "thor": lambda: handle_thor(account, pswd, url),
            "sta1": lambda: handle_sta1(account, pswd, url),
            "sta2": lambda: handle_sta2(account, pswd, url),
            "prod": lambda: handle_prod(account, pswd, url),
        },
    }

    handlers.get(website, {}).get(environment, default_handler)()


def default_handler():
    print("選擇的網站錯誤")


def handle_thor(account, pswd, url):
    open_url(url)
    login(account, pswd)


def handle_sta1(account, pswd, url):
    open_url(url)
    login(account, pswd)


def handle_sta2(account, pswd, url):
    open_url(url)
    login(account, pswd)


def handle_prod(account, pswd, url):
    open_url(url)
    login(account, pswd)


def open_url(url):
    global chromeWeb
    
    chromeWeb = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    chromeWeb.maximize_window()
    chromeWeb.get(url)


def login(account, pswd):
    # 取得驗證碼位置
    time.sleep(1)
    chromeWeb.save_screenshot("test.png")
    element = chromeWeb.find_element(By.ID, "verifyimg")
    verification_code = img(element)
    time.sleep(1)
    if verification_code != "111":
        while len(verification_code) != 4:
            chromeWeb.refresh()
            time.sleep(1)
            chromeWeb.get_screenshot_as_file("test.png")
            element = chromeWeb.find_element(By.ID, "verifyimg")
            verification_code = img(element)
    else:
        chromeWeb.quit()
        print("驗證碼解析失敗，請重新執行")
        return False
    # 登入開始
    chromeWeb.find_element(By.NAME, "UserName").send_keys(account)
    time.sleep(1)
    chromeWeb.find_element(By.NAME, "Password").send_keys(pswd)
    time.sleep(1)
    chromeWeb.find_element(By.XPATH, '//*[@id="txtInvalidation"]').send_keys(
        verification_code
    )
    print("輸入驗證碼完成")
    time.sleep(1)
    chromeWeb.find_element(By.NAME, "Submit").click()  # 登入按鈕
    return True


# 驗證碼解析用
def img(element):
    try:
        element.location  # 取得圖片位置
        element.size  # 取得高度寬度
        left = element.location["x"]  # 取得左邊
        right = element.location["x"] + element.size["width"]  # 取得右邊
        top = element.location["y"]  # 取得上面
        bottom = element.location["y"] + element.size["height"]  # 取得下邊

        # 切割出驗證碼
        img_code = Image.open(".\\test.png")
        img_code.load()  # 確保圖片完全載入
        img_code = img_code.crop((left, top, right, bottom))  # 順序一定要是 x,y,x+w,y+h
        img_code.save("verify.png", "png")

        # 對驗證碼進行處理
        img_code = img_code.convert("L")
        pix = img_code.load()
        w, h = img_code.size
        threshold = 205  # 畫素閾值

        # 遍歷所有畫素，大於閾值的為黑色
        for y in range(h):
            for x in range(w):
                if pix[x, y] < threshold:
                    pix[x, y] = 0
                else:
                    pix[x, y] = 255

        # 根據畫素二值結果重新生成圖片
        data = img_code.getdata()
        w, h = img_code.size
        black_point = 0
        for x in range(1, w - 1):
            for y in range(1, h - 1):
                mid_pixel = data[w * y + x]
                if mid_pixel < 50:
                    top_pixel = data[w * (y - 1) + x]
                    left_pixel = data[w * y + (x - 1)]
                    down_pixel = data[w * (y + 1) + x]
                    right_pixel = data[w * y + (x + 1)]
                    if top_pixel < 10:
                        black_point += 1
                    if left_pixel < 10:
                        black_point += 1
                    if down_pixel < 10:
                        black_point += 1
                    if right_pixel < 10:
                        black_point += 1
                    if black_point < 1:
                        img_code.putpixel((x, y), 255)
                    black_point = 0

        result = pytesseract.image_to_string(img_code)
        # 可能存在異常符號，用正則提取其中的數字
        regex = "\d+"
        result = "".join(re.findall(regex, result))

        os.remove("verify.png")
        os.remove("test.png")
    except Exception as e:
        print(e)
        return "111"

    return result


def oddsConversion(requset):
    return render(requset, "oddsConversion.html")


if __name__ == "__main__":
    upload()
