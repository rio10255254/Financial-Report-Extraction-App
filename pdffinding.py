from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import os
import re

# 初始化 WebDriver
firefox_options = Options()
firefox_options.headless = True  # 無頭模式，避免打開瀏覽器窗口
webdriver_service = Service(r"C:\Users\Sean Kang\Downloads\geckodriver-v0.34.0-win32\geckodriver.exe")  # geckodriver 的路徑，根據您的實際安裝進行調整

firefox_options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"

def fetch_pdf_links_liteon(url):
    pdf_links = []  # 存儲 PDF 連結的清單
    driver = None

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)

        # 等待頁面加載完成
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'svg-hover-outer clickable-row')]")))

        # 找到所有包含 PDF 連結的 tr 元素
        tr_elements = driver.find_elements(By.XPATH, "//tr[contains(@class, 'svg-hover-outer clickable-row')]")

        for tr in tr_elements:
            # 獲取 data-href 屬性，即 PDF 文件的鏈接
            pdf_link = tr.get_attribute('data-href')
            if pdf_link:
                pdf_links.append(pdf_link)

        # 打印找到的 PDF 連結及其名稱
        for link in pdf_links:
            # 從連結中獲取 PDF 檔案名稱
            file_name = link.split('/')[-1]
            print(f"Found downloadable PDF: {file_name}")

    except WebDriverException as e:
        print(f"Error accessing {url} with Selenium: {e}")

    finally:
        if driver:
            driver.quit()

# 執行函數，替換 URL 為您要查看的網頁
fetch_pdf_links_liteon("https://www.liteon.com/zh-tw/download/quarterly-reports/index/2023")
