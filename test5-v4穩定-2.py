#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# test5-v4.py
# @Author :  ()
# @Link   : 
# @Date   : 2024/6/25 上午11:14:58
# 完成Liteon, Acer

import os  # 匯入作業系統模組，用於處理文件和目錄操作
import time  # 匯入時間模組，用於處理延遲等時間相關操作
import requests  # 匯入requests模組，用於發送HTTP請求
import re  # 匯入正則表達式模組，用於字符串模式匹配
import threading  # 匯入多執行緒模組，用於並行執行任務
import tkinter as tk
from tkinter import Tk, Label, Entry, Button, END  # 匯入Tkinter模組，用於創建圖形用戶界面
from tkinter.scrolledtext import ScrolledText  # 匯入ScrolledText，用於可滾動的文本框
from queue import Queue  # 匯入佇列模組，用於執行緒之間的數據交換
from bs4 import BeautifulSoup  # 匯入BeautifulSoup，用於解析HTML和XML
from selenium import webdriver  # 匯入Selenium的webdriver，用於自動化網頁操作
from selenium.webdriver.firefox.service import Service as FirefoxService  # 匯入FirefoxService，用於設定Firefox瀏覽器服務
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.options import Options as FirefoxOptions  # 匯入FirefoxOptions，用於設定Firefox瀏覽器選項
from selenium.webdriver.common.by import By  # 匯入By模組，用於定位網頁元素
from selenium.webdriver.support.ui import WebDriverWait  # 匯入WebDriverWait，用於顯式等待
from selenium.webdriver.support import expected_conditions as EC  # 匯入預期條件模組，用於顯式等待的條件設置
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException  # 匯入Selenium的例外處理，用於處理各種異常情況
import urllib3  # 匯入urllib3，用於處理HTTP請求
import urllib.parse  # 匯入urllib.parse，用於URL解析
from concurrent.futures import ThreadPoolExecutor, as_completed  # 匯入ThreadPoolExecutor和as_completed，用於多執行緒並行處理
from queue import Queue, Empty  # 匯入Queue和Empty，用於佇列操作
import random  # 匯入隨機數模組，用於生成隨機數
from threading import Lock  # 匯入Lock，用於多執行緒的同步控制
from tkinter import Tk, Label, Entry, Button, END, Frame  # 再次匯入Tkinter模組的其它組件
import warnings  # 匯入警告模組，用於處理警告
import ttkbootstrap as ttk  # 匯入ttkbootstrap，用於改進Tkinter的樣式
from ttkbootstrap.constants import *  # 匯入ttkbootstrap的常量
from tkinter.scrolledtext import ScrolledText  # 再次匯入ScrolledText，用於可滾動的文本框
import threading  # 再次匯入多執行緒模組
from queue import Queue, Empty  # 再次匯入Queue和Empty，用於佇列操作
from concurrent.futures import ThreadPoolExecutor, as_completed  # 再次匯入ThreadPoolExecutor和as_completed，用於多執行緒並行處理
import os  # 再次匯入作業系統模組
import requests  # 再次匯入requests模組
from bs4 import BeautifulSoup  # 再次匯入BeautifulSoup
from tkinter import font  # 匯入Tkinter的font模組，用於設置字體
from tkinter import messagebox
from tkinter import ttk, scrolledtext, font, Tk
from tkinter import END, Tk
from urllib.parse import urljoin
from requests.exceptions import RequestException

# 修改ThreadPoolExecutor的最大工作數量，將其設置為更小的值
MAX_WORKERS = 3

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 要檢查的網站URL列表
site_urls = {
    # 無法成功下載之公司
    'Broadcom': 'https://investors.broadcom.com/financial-information/quarterly-results',
    'Micron': 'https://investors.micron.com/quarterly-results',
    'Quanta': 'https://www.quantatw.com/Quanta/chinese/investment/financials_qr3.aspx',
    'Celestica': 'https://corporate.celestica.com/quarterly-results',
    'Realtek': 'https://www.realtek.com/InvestorRelations/FinancialStatements?menu_id=461&lang=zh-TW', # 正在
    'TI': 'https://investor.ti.com/financial-information/earnings-annual-reports',
    'ST Micro': 'https://investors.st.com/financial-information/quarterly-results',
    'ADI': 'https://investor.analog.com/financial-info/quarterly-results',
    'Renesas': 'https://www.renesas.com/us/en/about/investor-relations/event/presentation/2023#financial_results',
    'Onsemi': 'https://investor.onsemi.com/financials',
    'Vishay': 'https://ir.vishay.com/financial-information/quarterly-results',
    'Qorvo': 'https://ir.qorvo.com/press-releases',
    'Diodes': 'https://investor.diodes.com/financial-performance-1',
    # 無法成功下載之公司尾端

    #可成功下載之公司
    'Acer': 'https://www.acer.com/corporate/zh/investor-relations/financials/annual-reports',
    'Nanya': 'https://www.nanya.com/tw/IR/39',
    'Liteon': 'https://www.liteon.com/zh-tw/download/quarterly-reports/index/2023',
    'MPS': 'https://www.monolithicpower.com/en/about-mps/investor-relations/press-releases.html',
    'AOS': 'https://investor.aosmd.com/financial-information/annual-reports-and-proxy/default.aspx',
    'SK hynix': 'https://www.skhynix.com/ir/UI-FR-IR12_T4',
    'Inventec': 'https://www.inventec.com/tw/finance-2',
    'NXP': 'https://investors.nxp.com/financial-information/financial-information-0',
    'Qualcomm': 'https://investor.qualcomm.com/financial-information/historical-financial-results',
    'AMD': 'https://ir.amd.com/financial-information/historical-financials',
    'HonHai': 'https://www.foxconn.com/zh-tw/investor-relations/financial-information/reports?category=quarterly',
    'Compal': 'https://www.compal.com/investor-relations/financial-release/#consolidated-financial',
    'Pegatron': 'https://www.pegatroncorp.com/investorRelation/downloadLibrary/type/financial_information/year/all',
    'Wistron': 'https://www.wistron.com/ch/Investors/Financials/FinancialReports',
    'Wiwynn': 'https://www.wiwynn.com/zh/investors#quarterlyresults',
    'Benchmark': 'https://ir.bench.com/financials/quarterly-results/default.aspx',
    'Flex': 'https://investors.flex.com/financials/quarterly-results/default.aspx',
    'Jabil':'https://investors.jabil.com/financials/quarterly-results/default.aspx',
    'Plexus':'https://investor.plexus.com/financials/quarterly-results/default.aspx',
    'Samina': 'https://ir.sanmina.com/financials/quarterly-results/default.aspx',
    'Delta': 'https://www.deltaww.com/en-US/Investors/financial-Reports',
    'Chicony': 'https://www.chicony.com/chicony/en/investors/investors/QuarterlyPresentations',
    'AcBel': 'https://www.acbel.com.tw/financial-reports',
    'MSI': 'https://www.msi.com/about/investor/financialInformation',
    'Giga-Byte': 'https://www.gigabyte.com/FileUpload/TW/SiteMap/134/index.html#tab134-2',
    'Asus': 'https://www.asus.com/tw/pages/investor/',
    'Intel': 'https://www.intc.com/financial-info/financial-results',
    'Nvidia': 'https://investor.nvidia.com/financial-info/financial-reports/default.aspx',
    'Samsung': 'https://www.samsung.com/global/ir/financial-information/audited-financial-statements/',
    'Kioxia': 'https://www.kioxia-holdings.com/en-jp/about/company.html#anc05',
    'Winbond': 'https://www.winbond.com/hq/about-winbond/investor/financial-information/financial-reports/?__locale=zh_TW',
    'Mediatek': 'https://corp.mediatek.tw/investor-relations/financial-information/quarterly-earnings/2023',
    'Infineon': 'https://www.infineon.com/cms/en/about-infineon/investor/reports-and-presentations/#financial-results',
    'Toshiba': 'https://www.global.toshiba/ww/ir/corporate/finance/er/er-list.html',
    'Ams-Osram': 'https://ams-osram.com/about-us/investor-relations/financial-results-and-reports',
    'Marvell': 'https://investor.marvell.com/quarterly-results',
    'Skyworks': 'https://investors.skyworksinc.com/annual-reports-and-proxies',
    'Rohm': 'https://www.rohm.com/ir/library/financial-report/', 
}

# 保存PDF文件的目錄
save_directory = 'C:\\PY-CH\\PY-CH-10'

# 檢查保存目錄是否存在,如果不存在則創建它
if not os.path.exists(save_directory):
    os.makedirs(save_directory)

# 設置目標年份和季度
target_year = '2023'  # 目標年份
target_keywords = ['Results', 'Release', 'finance-statement', '財務', 'Earnings Release', 'Financial Results', 'Quarterly Results', 'Annual Report', 'Consolicated Financial', 'Financial Statements']  # 目標關鍵字
exclude_keywords = ['逐字稿', '簡報', 'Presentation']  # 排除的關鍵字
no_filter_companies = ['Quanta', 'Broadcom']  # 不需要過濾的公司

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, Gecko) Chrome/58.0.3029.110 Safari/537.3'
}
# 設定HTTP請求的標頭，使其模仿從真實的瀏覽器發出的請求。這有助於避免某些網站的反爬蟲機制。

# 設置Firefox瀏覽器選項
firefox_options = FirefoxOptions()
firefox_options.add_argument("--headless")  # 設置無頭模式，這意味著瀏覽器在後台運行而不顯示圖形界面
firefox_options.add_argument("--disable-gpu")  # 禁用GPU硬件加速，通常在無頭模式下使用
firefox_options.add_argument("--window-size=1920,1080")  # 設置瀏覽器窗口的大小，這可以模擬高分辨率顯示器
firefox_options.add_argument("--start-maximized")  # 設置瀏覽器啟動時最大化窗口

# 指定geckodriver的路徑
webdriver_service = FirefoxService(r"C:\Users\Sean Kang\Downloads\geckodriver-v0.34.0-win32\geckodriver.exe")
# 設定geckodriver的路徑，這是Firefox的WebDriver，負責與Firefox瀏覽器進行交互。

firefox_options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

def show_developer_window():
    # 創建一個新的頂層窗口作為開發者選項窗口
    developer_window = tk.Toplevel()
    developer_window.title("Developer Options")  # 設定窗口標題為"Developer Options"
    developer_window.geometry("600x400")  # 設定窗口大小為600x400

    # 創建顯示現有公司名單及其URL的窗口
    company_list_frame = ttk.Frame(developer_window)
    company_list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # 設定框架填充並擴展，邊距為10

    # 創建一個帶滾動條的文字框來顯示公司名單和URL
    company_list_text = scrolledtext.ScrolledText(company_list_frame, wrap=tk.WORD)
    company_list_text.pack(fill=tk.BOTH, expand=True)  # 設定文字框填充並擴展
    company_list_text.insert(tk.END, "Current Site URLs:\n\n")  # 插入標題"Current Site URLs:"

    # 從site_urls字典中讀取公司和URL並插入到文字框中
    for company, url in site_urls.items():
        company_list_text.insert(tk.END, f"{company}\n{url}\n\n")  # 插入每個公司名和對應的URL

            
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

def convert_to_minguo_year(year):
    """將西元年轉換為民國年"""
    return str(int(year) - 1911)
    # 將傳入的西元年（公曆年）轉換為民國年（Minguo year）
    # 民國年 = 西元年 - 1911

def match_year_in_text(text, year):
    """檢查文本中是否包含指定年份的多種可能形式"""
    minguo_year = convert_to_minguo_year(year)  # 將西元年轉換為民國年
    short_year = year[-2:]  # 提取年份的後兩位數，例如2024年的短年份形式為24
    patterns = [
        year,  # 完整的西元年，例如2024
        minguo_year,  # 對應的民國年，例如113
        f'Q1{short_year}', f'Q2{short_year}', f'Q3{short_year}', f'Q4{short_year}',  # 季度表示，如Q124, Q224
        f'{short_year}Q1', f'{short_year}Q2', f'{short_year}Q3', f'{short_year}Q4',  # 季度表示，如24Q1, 24Q2
        f'{minguo_year}Q1', f'{minguo_year}Q2', f'{minguo_year}Q3', f'{minguo_year}Q4',  # 民國年的季度表示，如113Q1, 113Q2
        f'Q1{minguo_year}', f'Q2{minguo_year}', f'Q3{minguo_year}', f'Q4{minguo_year}',  # 民國年的季度表示，如Q1131, Q1132
        f'{year}Q1', f'{year}Q2', f'{year}Q3', f'{year}Q4',  # 完整年份的季度表示，如2023Q1, 2023Q2
        f'1Q{year}', f'2Q{year}', f'3Q{year}', f'4Q{year}',  # 季度年份表示，如1Q2023, 2Q2023
        f'Q{year[-2:]}1', f'Q{year[-2:]}2', f'Q{year[-2:]}3', f'Q{year[-2:]}4'  # 簡化年份的季度表示，如Q231, Q232
    ]
    return any(pat in text for pat in patterns)  # 檢查文本中是否包含上述任何一種年份表示形式

def match_keywords_in_text(text, keywords):
    """檢查文本中是否包含指定關鍵字"""
    return any(keyword.lower() in text.lower() for keyword in keywords)
    # 檢查文本是否包含任何指定的關鍵字（忽略大小寫）

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

def fetch_links_in_iframe(driver):
    pdf_links = []  # 初始化一個空列表，用於儲存PDF鏈接
    try:
        # 等待並獲取iframe中所有包含.pdf或download的鏈接元素
        pdf_links_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
        )
        # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
    except TimeoutException:
        pass  # 如果在指定時間內未找到PDF鏈接，則跳過
    return pdf_links  # 返回找到的PDF鏈接列表

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
def fetch_pdf_links_nxp(url, year, keywords, queue):
    queue.put(f"Fetching from URL: {url}\n")
    # 將當前正在抓取的URL放入佇列中，用於日誌記錄或後續處理

    response = requests.get(url, verify=False)
    # 發送HTTP GET請求以獲取指定URL的內容。verify=False表示忽略SSL證書驗證

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據

    file_links = soup.find_all('a', href=True)
    # 查找所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接

    pdf_links = []
    # 初始化一個空列表，用於存儲符合條件的PDF鏈接

    for link in file_links:
        # 遍歷所有找到的<a>標籤
        link_text = link.get_text(strip=True)
        # 獲取<a>標籤中的文本內容，並去除首尾的空白字符
        href = link['href']
        # 獲取<a>標籤的href屬性，即鏈接地址

        # queue.put(f"Checking link: {link_text}\n")
        # 將當前正在檢查的鏈接文本放入佇列中，用於日誌記錄或後續處理

        if any(keyword.lower() in link_text.lower() for keyword in keywords) and (year in link_text or year in href):
            # 檢查鏈接文本或href屬性中是否包含任何指定的關鍵字以及年份
            pdf_links.append(href if href.startswith('http') else 'https://investors.nxp.com' + href)
            # 如果鏈接是相對路徑，則將其轉換為絕對路徑，並添加到pdf_links列表中

            queue.put(f"Added link: {href}\n")
            # 將符合條件的鏈接放入佇列中，用於日誌記錄或後續處理

    return pdf_links
    # 返回找到的PDF鏈接列表

def is_pdf_nxp(url):
    """检查文件是否是PDF格式"""
    try:
        response = requests.head(url, allow_redirects=True, timeout=10)
        # 發送HTTP HEAD請求，允許重定向，設置超時時間為10秒
        content_type = response.headers.get('Content-Type')
        # 獲取響應標頭中的Content-Type字段
        return 'pdf' in content_type.lower()
        # 檢查Content-Type字段中是否包含'pdf'，如果包含則返回True，否則返回False
    except requests.RequestException as e:
        print(f"Error checking if URL is a PDF: {e}", url)
        # 捕獲HTTP請求異常，並打印錯誤信息
        return False
        # 返回False表示檢查失敗

def download_file_nxp(url, save_path):
    """下载文件并保存到指定路径"""
    try:
        response = requests.get(url, allow_redirects=True, timeout=30)
        # 發送HTTP GET請求下載文件，允許重定向，設置超時時間為30秒
        response.raise_for_status()
        # 如果響應狀態碼表示錯誤，則引發HTTPError異常
        if response.status_code == 200:
            # 如果響應狀態碼為200（成功）
            with open(save_path, 'wb') as f:
                f.write(response.content)
            # 打開指定的保存路徑，以二進制寫入模式保存文件內容
            return True
            # 返回True表示下載成功
    except requests.exceptions.RequestException as e:
        print(f"Failed to download {url}: {e}")
        # 捕獲HTTP請求異常，並打印錯誤信息
    return False
    # 返回False表示下載失敗

def download_pdfs_for_year_nxp(year, site_urls, keywords, queue, downloaded_files_box, root):
    """下載包含指定年份和關鍵字的財務報告PDF文件"""
    for site_name, site_url in site_urls.items():
        # 遍歷每個站點及其對應的URL
        queue.put(f"Processing site: {site_name}\n")
        # 將當前處理的站點名稱放入佇列，用於日誌記錄或後續處理
        file_urls = fetch_pdf_links_nxp(site_url, year, keywords, queue)
        # 調用函數獲取包含指定年份和關鍵字的PDF文件鏈接
        queue.put(f"Found PDF URLs for {site_name}: {file_urls}\n")
        # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
        if not file_urls:
            # 如果沒有找到PDF鏈接
            queue.put(f"No PDF links found for {site_name}.\n")
            # 將未找到PDF鏈接的信息放入佇列，用於日誌記錄或後續處理
        for file_url in file_urls:
            # 遍歷找到的每個PDF文件鏈接
            full_url = file_url if file_url.startswith('http') else f'https://{site_name.lower()}.com{file_url}'
            # 如果鏈接是相對路徑，將其轉換為絕對路徑
            file_name = f"{site_name}_{file_url.split('/')[-1]}"
            # 構建文件名，格式為"站點名稱_文件名"

            # 創建公司名稱的子文件夾
            company_directory = os.path.join(save_directory, site_name)
            # 構建公司目錄的路徑
            if not os.path.exists(company_directory):
                os.makedirs(company_directory)
                # 如果公司目錄不存在，則創建它

            save_path = os.path.join(company_directory, file_name)
            # 構建文件的保存路徑
            queue.put(f"Downloading file from: {full_url}\n")
            # 將下載文件的URL放入佇列，用於日誌記錄或後續處理
            if download_file_nxp(full_url, save_path):
                # 如果文件下載成功
                if is_pdf_nxp(full_url):
                    # 如果下載的文件是PDF格式
                    queue.put(f"Downloaded PDF: {file_name}\n")
                    # 將下載成功的PDF文件名稱放入佇列，用於日誌記錄或後續處理
                    # 插入已下載的文件名稱到輸出框
                    downloaded_files_box.insert(END, f"下載完成: {file_name}\n")
                    downloaded_files_box.see(END)
                    root.update_idletasks()
                else:
                    # 如果下載的文件不是PDF格式
                    queue.put(f"Skipped non-PDF file after download: {full_url}\n")
                    # 將跳過的非PDF文件信息放入佇列，用於日誌記錄或後續處理
                    os.remove(save_path)
                    # 刪除非PDF文件
            else:
                queue.put(f"Failed to download: {file_name}\n")
                # 如果文件下載失敗，將失敗信息放入佇列，用於日誌記錄或後續處理
                
def fetch_pdf_links(url, year, site_name, keywords):
    """從給定的URL中提取包含指定年份和關鍵字的PDF文件鏈接"""
    if site_name == 'NXP':
        return fetch_pdf_links_nxp(url, year, keywords)
        # 如果站點名稱是NXP，則調用fetch_pdf_links_nxp函數來提取PDF鏈接
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

def fetch_pdf_links_intel(url, queue):
    pdf_links = []  # 初始化一個空列表，用於儲存PDF鏈接
    driver = None  # 初始化driver變數
    try:
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        # 將當前URL和頁面標題添加到佇列中，供後續處理
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        # 滾動到頁面的底部，以確保所有內容都被加載
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)  # 等待5秒鐘以確保頁面加載完成
        
        try:
            # 等待並獲取所有包含.pdf或download的鏈接元素
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            # 如果在指定時間內未找到PDF鏈接，則發出警告
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        # 查找所有iframe元素並處理其中的鏈接
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = []
            while iframes:
                iframe = iframes.pop()
                driver.switch_to.frame(iframe)  # 切換到iframe
                futures.append(executor.submit(fetch_links_in_iframe, driver))  # 提交任務以獲取iframe中的鏈接
                driver.switch_to.default_content()  # 切換回主文檔

            for future in as_completed(futures):
                pdf_links.extend(future.result())  # 將每個future的結果添加到pdf_links中
                
    except WebDriverException as e:
        # 處理WebDriver異常，並將錯誤信息添加到佇列中
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        # 確保在結束時關閉瀏覽器
        if driver:
            driver.quit()

    return pdf_links  # 返回找到的PDF鏈接列表

def fetch_pdf_links_qualcomm(url, year, keywords, queue):
    # 函數從給定的URL中提取包含指定年份和關鍵字的PDF文件鏈接
    queue.put(f"Fetching from URL: {url}\n")
    # 將當前正在抓取的URL放入佇列，用於日誌記錄或後續處理

    response = requests.get(url)
    # 發送HTTP GET請求以獲取指定URL的內容

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據

    file_links = soup.find_all('a', href=True)
    # 查找所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接

    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    for link in file_links:
        # 遍歷所有找到的<a>標籤
        link_text = link.get_text(strip=True)
        # 獲取<a>標籤中的文本內容，並去除首尾的空白字符
        href = link['href']
        # 獲取<a>標籤的href屬性，即鏈接地址

        # queue.put(f"Checking link: {link_text}\n")
        # 將當前正在檢查的鏈接文本放入佇列，用於日誌記錄或後續處理

        if any(keyword.lower() in link_text.lower() for keyword in keywords) and (year in link_text or year in href):
            # 檢查鏈接文本或href屬性中是否包含任何指定的關鍵字以及年份
            pdf_link = href if href.startswith('http') else 'https://investor.qualcomm.com' + href
            # 如果鏈接是相對路徑，將其轉換為絕對路徑

            pdf_links.append((pdf_link, link_text))
            # 將符合條件的鏈接和文本添加到pdf_links列表中

            queue.put(f"Added link: {pdf_link}\n")
            # 將符合條件的鏈接放入佇列，用於日誌記錄或後續處理

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_amd(url, year, keywords, queue):
    # 函數從給定的URL中提取包含指定年份和關鍵字的PDF文件鏈接
    queue.put(f"Fetching from URL: {url}\n")
    # 將當前正在抓取的URL放入佇列，用於日誌記錄或後續處理

    response = requests.get(url)
    # 發送HTTP GET請求以獲取指定URL的內容

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據

    file_links = soup.find_all('a', href=True)
    # 查找所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接

    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    base_url = '/'.join(url.split('/')[:3])
    # 獲取基本URL，用於處理相對鏈接

    def fetch_detail_page(href):
        detail_url = href if href.startswith('http') else base_url + href
        # 如果鏈接是相對路徑，將其轉換為絕對路徑
        queue.put(f"Fetching details from: {detail_url}\n")
        # 將當前正在抓取的詳情頁URL放入佇列，用於日誌記錄或後續處理
        try:
            detail_response = requests.get(detail_url, allow_redirects=False, timeout=30)
            # 發送HTTP GET請求以獲取詳情頁的內容，設置超時時間為30秒，不允許重定向
            detail_response.raise_for_status()
            # 如果響應狀態碼表示錯誤，則引發HTTPError異常
            detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
            # 使用BeautifulSoup解析詳情頁的HTML內容
            detail_links = detail_soup.find_all('a', href=True)
            # 查找詳情頁中所有包含href屬性的<a>標籤
            for detail_link in detail_links:
                # 遍歷所有找到的<a>標籤
                detail_href = detail_link['href']
                # 獲取<a>標籤的href屬性，即鏈接地址
                if detail_href.endswith('.pdf'):
                    # 如果鏈接地址以.pdf結尾，表示這是一個PDF文件
                    pdf_link = detail_href if detail_href.startswith('http') else base_url + detail_href
                    # 如果鏈接是相對路徑，將其轉換為絕對路徑
                    queue.put(f"Found PDF link: {pdf_link}\n")
                    # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
                    return pdf_link, detail_link.get_text(strip=True)
                    # 返回PDF鏈接和鏈接文本
        except requests.exceptions.RequestException as e:
            queue.put(f"Error fetching details from {detail_url}: {e}\n")
            # 捕獲HTTP請求異常，並將錯誤信息放入佇列，用於日誌記錄或後續處理
        return None
        # 返回None表示未找到PDF鏈接

    with ThreadPoolExecutor(max_workers=10) as executor:
        # 使用線程池來並行處理詳情頁抓取任務
        future_to_url = {executor.submit(fetch_detail_page, link['href']): link['href'] for link in file_links if year in link.get_text(strip=True) or year in link['href']}
        # 提交詳情頁抓取任務到線程池，並將任務對應的URL存入字典
        for future in as_completed(future_to_url):
            # 當任務完成時，獲取結果
            result = future.result()
            if result:
                pdf_links.append(result)
                # 將找到的PDF鏈接添加到pdf_links列表中

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_honhai(url, queue):
    # 函數從指定的URL中提取PDF文件鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        # 將當前URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title: {driver.title}\n")
        # 將當前頁面標題放入佇列，用於日誌記錄或後續處理
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # 滾動到頁面的底部，以確保所有內容都被加載
        time.sleep(5)
        # 等待5秒鐘以確保頁面加載完成
        
        try:
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 等待並獲取所有包含.pdf或download的鏈接元素
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
            # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
            # 如果在指定時間內未找到PDF鏈接，則發出警告

        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        # 查找所有iframe元素
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            # 切換到iframe
            try:
                pdf_links_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
                )
                # 等待並獲取iframe中所有包含.pdf或download的鏈接元素
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
                # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
            except TimeoutException:
                pass
                # 如果在指定時間內未找到PDF鏈接，則跳過
            iframes.extend(driver.find_elements(By.TAG_NAME, 'iframe'))
            # 繼續查找嵌套的iframe
            driver.switch_to.default_content()
            # 切換回主文檔

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def manual_navigate_honhai(output_box, queue):
    # 手動導航並提取鴻海網站中的PDF鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get('https://www.foxconn.com/zh-tw')
        # 訪問鴻海網站的主頁面
        
        time.sleep(30)
        # 等待30秒，以確保頁面加載完成

        queue.put("Navigating to https://www.foxconn.com/zh-tw/investor-relations\n")
        # 將導航信息放入佇列，用於日誌記錄或後續處理
        investor_relations_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/zh-tw/investor-relations')]"))
        )
        # 等待並獲取投資者關係頁面的鏈接
        investor_relations_link.click()
        # 點擊投資者關係頁面的鏈接
        time.sleep(30)
        # 等待30秒，以確保頁面加載完成

        queue.put("Navigating to https://www.foxconn.com/zh-tw/investor-relations/financial-information/reports?category=quarterly\n")
        # 將導航信息放入佇列，用於日誌記錄或後續處理
        financial_reports_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/zh-tw/investor-relations/financial-information/reports?category=quarterly')]"))
        )
        # 等待並獲取財務報告頁面的鏈接
        financial_reports_link.click()
        # 點擊財務報告頁面的鏈接
        time.sleep(30)
        # 等待30秒，以確保頁面加載完成
        
        try:
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 等待並獲取所有包含.pdf或download的鏈接元素
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
            # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        except TimeoutException:
            queue.put("Warning: PDF links not found during manual navigation\n")
            # 如果在指定時間內未找到PDF鏈接，則發出警告
        
    except WebDriverException as e:
        queue.put(f"Error during manual navigation for Hon Hai: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_liteon(url, queue, company_name, save_directory):
    """從Liteon網站的特定URL中提取PDF文件連結"""
    pdf_links = []  # 用於存儲PDF連結的清單
    driver = None

    try:
        firefox_options = Options()
        firefox_options.headless = True
        webdriver_service = FirefoxService()

        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)  # 打開指定的URL

        # 等待表格內容加載完成
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//tbody")))

        # 找到所有包含 PDF 連結的 tr 元素
        tr_elements = driver.find_elements(By.XPATH, "//tr[contains(@class, 'svg-hover-outer')]")

        for tr in tr_elements:
            try:
                # 獲取 data-href 屬性的值，即 PDF 下載連結
                pdf_link = tr.get_attribute('data-href')
                if pdf_link and pdf_link.endswith('.pdf'):
                    full_pdf_link = f"https://www.liteon.com/{pdf_link}"
                    # 使用 tr 中的文字來生成鏈接文本
                    link_text = tr.text.strip()
                    pdf_links.append((full_pdf_link, link_text))
            except NoSuchElementException as e:
                queue.put(f"No PDF link found in row: {tr.text.strip()}, {e}")

        if pdf_links:
            for link, text in pdf_links:
                queue.put(f"Processing link: {link} with text: {text}\n")  # 將正在處理的連結放入隊列中

        # 調用過濾函式來確保連結的格式符合預期
        return pdf_links

    except Exception as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")  # 如果訪問URL時發生異常，將錯誤信息放入隊列中
        return []

    finally:
        if driver:
            driver.quit()  # 關閉WebDriver

def fetch_pdf_links_chicony(queue):
    # 函數從群光網站中提取PDF文件鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get('https://www.chicony.com/chicony/en/investors/investors/QuarterlyPresentations')
        # 訪問群光網站的季度報告頁面
        
        queue.put(f"Current URL: {driver.current_url}\n")
        # 將當前URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title: {driver.title}\n")
        # 將當前頁面標題放入佇列，用於日誌記錄或後續處理
        
        time.sleep(10)
        # 等待10秒，以確保頁面加載完成
        
        pdf_links_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
        )
        # 等待並獲取所有包含.pdf或download的鏈接元素
        pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        
        queue.put(f"Chicony found links: {pdf_links}\n")
        # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
        
    except WebDriverException as e:
        queue.put(f"Error accessing Chicony site with Selenium: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_acer(queue):
    # 函數從宏碁網站中提取PDF文件鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get('https://www.acer.com/corporate/zh/investor-relations/financials/annual-reports')
        # 訪問宏碁網站的年度報告頁面
        
        queue.put(f"Current URL: {driver.current_url}\n")
        # 將當前URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title: {driver.title}\n")
        # 將當前頁面標題放入佇列，用於日誌記錄或後續處理
        
        time.sleep(20)
        # 等待20秒，以確保頁面加載完成
        
        pdf_links_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
        )
        # 等待並獲取所有包含.pdf或download的鏈接元素
        pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        
        queue.put(f"Acer found links: {pdf_links}\n")
        # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
        
    except TimeoutException:
        queue.put("TimeoutException: Page took too long to load or PDF links not found\n")
        # 捕獲超時異常，並將錯誤信息放入佇列

    except WebDriverException as e:
        queue.put(f"Error accessing Acer site with Selenium: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_pegatron(url, queue):
    # 函數從和碩網站中提取PDF文件鏈接
    try:
        response = requests.get(url, headers=headers, allow_redirects=True)
        # 發送HTTP GET請求以獲取指定URL的內容，允許重定向
        response.raise_for_status()
        # 如果響應狀態碼表示錯誤，則引發HTTPError異常
    except requests.exceptions.RequestException as e:
        queue.put(f"Error accessing {url}: {e}\n")
        # 捕獲HTTP請求異常，並將錯誤信息放入佇列
        return []
        # 返回空列表表示未找到PDF鏈接

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據
    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    for a_tag in soup.find_all('a', href=True):
        # 遍歷所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接
        href = a_tag['href']
        # 獲取<a>標籤的href屬性，即鏈接地址
        full_url = href if href.startswith('http') else requests.compat.urljoin(url, href)
        # 如果鏈接是相對路徑，將其轉換為絕對路徑
        pdf_links.append((full_url, a_tag.get_text(strip=True)))
        # 將符合條件的鏈接和文本添加到pdf_links列表中

    queue.put(f"Pegatron found links: {pdf_links}\n")
    # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_compal(url, queue):
    # 函數從仁寶網站中提取PDF文件鏈接
    try:
        response = requests.get(url, headers=headers, allow_redirects=True)
        # 發送HTTP GET請求以獲取指定URL的內容，允許重定向
        response.raise_for_status()
        # 如果響應狀態碼表示錯誤，則引發HTTPError異常
    except requests.exceptions.RequestException as e:
        queue.put(f"Error accessing {url}: {e}\n")
        # 捕獲HTTP請求異常，並將錯誤信息放入佇列
        return []
        # 返回空列表表示未找到PDF鏈接

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據
    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    for a_tag in soup.find_all('a', href=True):
        # 遍歷所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接
        href = a_tag['href']
        # 獲取<a>標籤的href屬性，即鏈接地址
        if href.startswith('javascript'):
            continue
            # 跳過以javascript開頭的鏈接
        full_url = href if href.startswith('http') else requests.compat.urljoin(url, href)
        # 如果鏈接是相對路徑，將其轉換為絕對路徑
        pdf_links.append((full_url, a_tag.get_text(strip=True)))
        # 將符合條件的鏈接和文本添加到pdf_links列表中

    queue.put(f"Compal found links: {pdf_links}\n")
    # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_wistron(url, queue):
    # 函數從緯創網站中提取PDF文件鏈接
    try:
        response = requests.get(url, headers=headers, allow_redirects=True)
        # 發送HTTP GET請求以獲取指定URL的內容，允許重定向
        response.raise_for_status()
        # 如果響應狀態碼表示錯誤，則引發HTTPError異常
    except requests.exceptions.RequestException as e:
        queue.put(f"Error accessing {url}: {e}\n")
        # 捕獲HTTP請求異常，並將錯誤信息放入佇列
        return []
        # 返回空列表表示未找到PDF鏈接

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據
    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    for a_tag in soup.find_all('a', href=True):
        # 遍歷所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接
        href = a_tag['href']
        # 獲取<a>標籤的href屬性，即鏈接地址
        if href.startswith('javascript'):
            continue
            # 跳過以javascript開頭的鏈接
        full_url = href if href.startswith('http') else requests.compat.urljoin(url, href)
        # 如果鏈接是相對路徑，將其轉換為絕對路徑
        pdf_links.append((full_url, a_tag.get_text(strip=True)))
        # 將符合條件的鏈接和文本添加到pdf_links列表中

    queue.put(f"Wistron found links: {pdf_links}\n")
    # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_inventec(url, queue, target_year, downloaded_files_box, root):
    """從英業達網站中提取PDF文件鏈接，處理具有年份選單的情況"""
    pdf_links = []
    driver = None

    def convert_to_minguo_year(year):
        """將西元年轉換為民國年"""
        return str(int(year) - 1911) + "年"

    minguo_year = convert_to_minguo_year(target_year)

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)

        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)

        try:
            year_dropdown = None

            try:
                year_dropdown = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//select[@id='year-select']"))
                )
                queue.put(f"Found year dropdown by ID\n")
            except TimeoutException:
                queue.put(f"Year dropdown not found by ID, trying alternative methods\n")

            if not year_dropdown:
                try:
                    year_dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'year')]"))
                    )
                    queue.put(f"Found year dropdown by class\n")
                except TimeoutException:
                    queue.put(f"Year dropdown not found by class, trying generic select tag\n")

            if not year_dropdown:
                try:
                    year_dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//select"))
                    )
                    queue.put(f"Found year dropdown by generic select tag\n")
                except TimeoutException:
                    queue.put(f"Generic select tag not found, listing all select elements\n")

            if not year_dropdown:
                select_elements = driver.find_elements(By.TAG_NAME, 'select')
                for select_element in select_elements:
                    queue.put(f"Select element found: {select_element.get_attribute('outerHTML')}\n")

            if year_dropdown:
                driver.execute_script("arguments[0].scrollIntoView(true);", year_dropdown)
                time.sleep(2)
                driver.execute_script("arguments[0].style.display = 'block';", year_dropdown)
                options = year_dropdown.find_elements(By.TAG_NAME, 'option')
                for option in options:
                    queue.put(f"Option: {option.text}\n")
                target_option = None
                for option in options:
                    if option.text == target_year or option.text == minguo_year:
                        target_option = option
                        break
                if target_option:
                    driver.execute_script("arguments[0].scrollIntoView(true);", target_option)
                    time.sleep(2)
                    target_option.click()
                    queue.put(f"Clicked year option: {target_option.text}\n")

                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", year_dropdown)
                    time.sleep(5)

                    queue.put("New content loaded successfully\n")
                else:
                    queue.put(f"Year option {target_year}/{minguo_year} not found in dropdown\n")
            else:
                queue.put(f"Year dropdown or target year {target_year}/{minguo_year} not found on page: {url}\n")

        except TimeoutException:
            queue.put(f"Warning: Year dropdown or target year {target_year}/{minguo_year} not found on page: {url}\n")

        try:
            scroll_height = driver.execute_script("return document.body.scrollHeight")
            for _ in range(10):  # 假設最多滾動10次
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(3)
                pdf_buttons = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')] | //a[contains(text(), '下載')]")
                for button in pdf_buttons:
                    href = button.get_attribute('href')
                    if not href:
                        button.click()
                        time.sleep(3)
                        pdf_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]")
                        for link in pdf_links_elements:
                            href = link.get_attribute('href')
                            if href and href not in pdf_links:
                                pdf_links.append((href, link.text.strip()))
                                queue.put(f"Found PDF link: {href}\n")
                                # 生成公司目录
                                company_name = "Inventec"
                                company_dir = os.path.join("C:\\PY-CH\\PY-CH-8", company_name)
                                if not os.path.exists(company_dir):
                                    os.makedirs(company_dir)
                                file_name = f"{company_name}_{sanitize_filename(href.split('/')[-1])}"
                                save_path = os.path.join(company_dir, file_name)
                                if download_file(href, save_path):
                                    downloaded_files_box.insert(END, f"下載完成: {file_name}\n")
                                    downloaded_files_box.see(END)
                                    root.update_idletasks()
                                    return pdf_links  # 一旦成功下載，就返回並結束函數
                    else:
                        pdf_links.append((href, button.text.strip()))
                        queue.put(f"Found PDF link: {href}\n")
                        # 生成公司目录
                        company_name = "Inventec"
                        company_dir = os.path.join("C:\\PY-CH\\PY-CH-8", company_name)
                        if not os.path.exists(company_dir):
                            os.makedirs(company_dir)
                        file_name = f"{company_name}_{sanitize_filename(href.split('/')[-1])}"
                        save_path = os.path.join(company_dir, file_name)
                        if download_file(href, save_path):
                            downloaded_files_box.insert(END, f"下載完成: {file_name}\n")
                            downloaded_files_box.see(END)
                            root.update_idletasks()
                            return pdf_links  # 一旦成功下載，就返回並結束函數
                if driver.execute_script("return document.body.scrollHeight") == scroll_height:
                    break  # 如果頁面高度沒有變化，停止滾動
                scroll_height = driver.execute_script("return document.body.scrollHeight")

            queue.put(f"Found PDF links: {[link for link, text in pdf_links]}\n")
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page after selecting year: {url}\n")

        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            iframes.extend(driver.find_elements(By.TAG_NAME, 'iframe'))
            driver.switch_to.default_content()

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")

    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_quanta(url, queue):
    # 函數從廣達網站中提取PDF文件鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get(url)
        # 訪問廣達網站的指定頁面
        
        queue.put(f"Current URL: {driver.current_url}\n")
        # 將當前URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title: {driver.title}\n")
        # 將當前頁面標題放入佇列，用於日誌記錄或後續處理
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # 滾動到頁面的底部，以確保所有內容都被加載
        time.sleep(5)
        # 等待5秒，以確保頁面加載完成
        
        try:
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 等待並獲取所有包含.pdf或download的鏈接元素
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
            # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
            # 如果在指定時間內未找到PDF鏈接，則發出警告

        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        # 查找所有iframe元素
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            # 切換到iframe
            try:
                pdf_links_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
                )
                # 等待並獲取iframe中所有包含.pdf或download的鏈接元素
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
                # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
            except TimeoutException:
                pass
                # 如果在指定時間內未找到PDF鏈接，則跳過
            iframes.extend(driver.find_elements(By.TAG_NAME, 'iframe'))
            # 繼續查找嵌套的iframe
            driver.switch_to.default_content()
            # 切換回主文檔

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def manual_navigate_quanta(output_box, queue):
    # 手動導航並提取廣達網站中的PDF鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get('https://www.quantatw.com/Quanta/chinese/index.aspx')
        # 訪問廣達網站的主頁面
        
        queue.put(f"Current URL: {driver.current_url}\n")
        # 將當前URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title: {driver.title}\n")
        # 將當前頁面標題放入佇列，用於日誌記錄或後續處理
        
        time.sleep(60)
        # 等待60秒，以確保頁面加載完成

        queue.put("Navigating to https://www.quantatw.com/Quanta/chinese/investment/financials_ms.aspx\n")
        # 將導航信息放入佇列，用於日誌記錄或後續處理
        financials_ms_link = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '財務月報表')]"))
        )
        # 等待並獲取財務月報表頁面的鏈接
        financials_ms_link.click()
        # 點擊財務月報表頁面的鏈接
        time.sleep(60)
        # 等待60秒，以確保頁面加載完成

        queue.put(f"Current URL after navigation: {driver.current_url}\n")
        # 將導航後的URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title after navigation: {driver.title}\n")
        # 將導航後的頁面標題放入佇列，用於日誌記錄或後續處理
        queue.put("Navigating to https://www.quantatw.com/Quanta/chinese/investment/financials_qr3.aspx\n")
        # 將導航信息放入佇列，用於日誌記錄或後續處理
        financials_qr3_link = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '財務季報表')]"))
        )
        # 等待並獲取財務季報表頁面的鏈接
        financials_qr3_link.click()
        # 點擊財務季報表頁面的鏈接
        time.sleep(60)
        # 等待60秒，以確保頁面加載完成

        queue.put(f"Current URL after navigation: {driver.current_url}\n")
        # 將導航後的URL放入佇列，用於日誌記錄或後續處理
        queue.put(f"Page title after navigation: {driver.title}\n")
        # 將導航後的頁面標題放入佇列，用於日誌記錄或後續處理
        
        try:
            pdf_links_elements = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 等待並獲取所有包含.pdf或download的鏈接元素
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
            # 提取這些鏈接的href屬性和文本，並添加到pdf_links列表中
        except TimeoutException:
            queue.put("Warning: PDF links not found during manual navigation\n")
            # 如果在指定時間內未找到PDF鏈接，則發出警告
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # 滾動到頁面的底部，以確保所有內容都被加載
        time.sleep(5)
        # 等待5秒，以確保頁面加載完成
        
    except WebDriverException as e:
        queue.put(f"Error during manual navigation for Quanta: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

def fetch_pdf_links_wiwynn(url, queue):
    # 函數從緯穎網站中提取PDF文件鏈接
    try:
        response = requests.get(url, headers=headers, allow_redirects=True)
        # 發送HTTP GET請求以獲取指定URL的內容，允許重定向
        response.raise_for_status()
        # 如果響應狀態碼表示錯誤，則引發HTTPError異常
    except requests.exceptions.RequestException as e:
        queue.put(f"Error accessing {url}: {e}\n")
        # 捕獲HTTP請求異常，並將錯誤信息放入佇列
        return []
        # 返回空列表表示未找到PDF鏈接

    soup = BeautifulSoup(response.text, 'html.parser')
    # 使用BeautifulSoup解析HTML內容，以便後續提取所需數據
    pdf_links = []
    # 初始化一個空列表，用於儲存符合條件的PDF鏈接

    for a_tag in soup.find_all('a', href=True):
        # 遍歷所有包含href屬性的<a>標籤，這些標籤可能包含PDF文件的鏈接
        href = a_tag['href']
        # 獲取<a>標籤的href屬性，即鏈接地址
        if href.startswith('javascript'):
            continue
            # 跳過以javascript開頭的鏈接
        full_url = href if href.startswith('http') else requests.compat.urljoin(url, href)
        # 如果鏈接是相對路徑，將其轉換為絕對路徑
        pdf_links.append((full_url, a_tag.get_text(strip=True)))
        # 將符合條件的鏈接和文本添加到pdf_links列表中

    queue.put(f"Wiwynn found links: {pdf_links}\n")
    # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
    return pdf_links
    # 返回找到的PDF鏈接列表

# def fetch_pdf_links_micron(url, queue):
#     driver = None
#     pdf_links = set()
#     try:
#         driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
#         driver.get(url)
#         queue.put(f"Current URL: {driver.current_url}\n")
#         queue.put(f"Page title: {driver.title}\n")

#         # 确保页面加载完成
#         time.sleep(10)
#         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         time.sleep(5)

#         # 获取页面上所有的链接
#         links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
#         for link in links:
#             href = link.get_attribute('href')
#             if href and ('.pdf' in href or 'download' in href):
#                 pdf_links.add(href)
#                 queue.put(f"Found PDF link: {href}\n")
        
#         # 检查页面中所有的 iframe
#         iframes = driver.find_elements(By.TAG_NAME, 'iframe')
#         for iframe in iframes:
#             driver.switch_to.frame(iframe)
#             links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
#             for link in links:
#                 href = link.get_attribute('href')
#                 if href and ('.pdf' in href or 'download' in href):
#                     pdf_links.add(href)
#                     queue.put(f"Found PDF link in iframe: {href}\n")
#             driver.switch_to.default_content()

#         # 处理动态加载的内容，尝试点击“Load More”按钮或其他动态加载的元素
#         while True:
#             try:
#                 load_more_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Load More')]")
#                 driver.execute_script("arguments[0].click();", load_more_button)
#                 time.sleep(5)

#                 # 再次获取页面上所有的链接
#                 links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
#                 for link in links:
#                     href = link.get_attribute('href')
#                     if href and ('.pdf' in href or 'download' in href):
#                         pdf_links.add(href)
#                         queue.put(f"Found PDF link: {href}\n")
#             except:
#                 break

#         # 再次检查页面上的所有链接，确保没有遗漏
#         links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
#         for link in links:
#             href = link.get_attribute('href')
#             if href and ('.pdf' in href or 'download' in href):
#                 pdf_links.add(href)
#                 queue.put(f"Found PDF link: {href}\n")

#     except WebDriverException as e:
#         queue.put(f"Error accessing {url} with Selenium: {e}\n")
#     finally:
#         if driver:
#             driver.quit()

#     return list(pdf_links)

def fetch_pdf_links_micron(url, queue):
    # time out
    pdf_links = []  # 存储PDF链接的列表
    retries = 3  # 设置最大重试次数
    timeout = 20  # 设置超时时间为20秒

    for attempt in range(retries):
        try:
            # 发送 GET 请求获取网页内容，设置超时时间
            response = requests.get(url, timeout=timeout)
            response.raise_for_status()  # 检查请求是否成功

            # 使用 BeautifulSoup 解析 HTML
            soup = BeautifulSoup(response.content, 'html.parser')

            # 找到所有包含 PDF 下载链接的 <a> 标签
            for link in soup.find_all('a', href=True):
                href = link['href']
                # 检查 <a> 标签的 type 属性是否包含 'application/pdf'
                if 'type' in link.attrs and 'pdf' in link['type']:
                    full_pdf_link = urljoin("https://investors.micron.com", href)
                    link_text = link.get_text(strip=True)
                    pdf_links.append((full_pdf_link, link_text))
                    queue.put(f"Found PDF link: {full_pdf_link} with text: {link_text}\n")  # 将找到的PDF链接放入队列中

            # 如果成功获取了 PDF 链接，则跳出循环
            break

        except requests.RequestException as e:
            if attempt < retries - 1:
                # 输出重试信息
                print(f"Attempt {attempt + 1} failed: {e}. Retrying...")
                time.sleep(5)  # 等待5秒后重试
            else:
                queue.put(f"Error accessing URL {url}: {e}\n")  # 如果超过重试次数仍然失败，将错误信息放入队列中
                break

    return pdf_links

def fetch_pdf_links_nanya(url, year, keywords, queue):
    queue.put(f"Fetching from URL: {url}\n")  # 將正在抓取的URL放入隊列中
    base_url = '/'.join(url.split('/')[:3])  # 提取基本URL
    visited_urls = set()  # 記錄已訪問的URL集合
    queue_urls = [url]  # 待處理的URL隊列
    pdf_links = []  # 用於存儲PDF連結的清單

    while queue_urls:
        current_url = queue_urls.pop(0)  # 從隊列中取出當前處理的URL
        if current_url in visited_urls:
            continue  # 如果URL已被訪問過，則跳過
        visited_urls.add(current_url)  # 將當前URL加入已訪問集合

        try:
            response = requests.get(current_url, timeout=30)  # 發送GET請求
            response.raise_for_status()  # 檢查請求是否成功
            soup = BeautifulSoup(response.text, 'html.parser')  # 解析HTML內容
            links = soup.find_all('a', href=True)  # 找到所有包含href屬性的<a>標籤
            
            for link in links:
                href = link['href']  # 獲取連結的href屬性
                full_url = href if href.startswith('http') else base_url + href  # 將相對連結轉換為完整URL

                # 如果連結或連結文字中包含任何關鍵字且包含年份，則認為是PDF連結
                if any(keyword.lower() in href.lower() for keyword in keywords) and year in href:
                    pdf_links.append((full_url, link.get_text(strip=True)))  # 將連結和文字添加到pdf_links清單中
                    queue.put(f"Found PDF link: {full_url}\n")  # 將找到的PDF連結放入隊列中
                elif any(keyword.lower() in link.get_text(strip=True).lower() for keyword in keywords) and year in link.get_text(strip=True):
                    pdf_links.append((full_url, link.get_text(strip=True)))  # 將連結和文字添加到pdf_links清單中
                    queue.put(f"Found PDF link: {full_url}\n")  # 將找到的PDF連結放入隊列中
                # 如果連結中包含與年份、季度或財務相關的詞，則將URL加入隊列以進一步處理
                elif 'year' in href.lower() or 'quarter' in href.lower() or 'financial' in href.lower() or '財務' in href or '報告' in href:
                    queue_urls.append(full_url)  # 將符合條件的URL加入隊列
                    queue.put(f"Adding URL to queue: {full_url}\n")  # 將添加的URL放入隊列中

        except requests.RequestException as e:
            queue.put(f"Error fetching details from {current_url}: {e}\n")  # 如果請求發生異常，將錯誤信息放入隊列中

    return pdf_links  # 返回PDF連結清單

# def fetch_pdf_links_broadcom(url, queue):
#     pdf_links = []  # 用於存儲PDF連結的清單
#     driver = None  # 初始化WebDriver為None
#     try:
#         driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)  # 使用Firefox啟動WebDriver
#         driver.get(url)  # 打開指定的URL
        
#         queue.put(f"Current URL: {driver.current_url}\n")  # 將當前URL放入隊列中
#         queue.put(f"Page title: {driver.title}\n")  # 將頁面標題放入隊列中
        
#         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 滾動到頁面底部
#         time.sleep(10)  # 等待10秒以確保所有元素都加載完成
        
#         try:
#             pdf_links_elements = WebDriverWait(driver, 60).until(
#                 EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
#             )  # 等待並查找所有包含".pdf"連結的元素
#             pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]  # 獲取每個連結的href屬性和文字
#         except TimeoutException:
#             queue.put(f"Warning: PDF links not found on page: {url}\n")  # 如果超時未找到PDF連結，將警告信息放入隊列中
        
#         iframes = driver.find_elements(By.TAG_NAME, 'iframe')  # 查找所有iframe元素
#         while iframes:
#             iframe = iframes.pop()  # 取出最後一個iframe
#             driver.switch_to.frame(iframe)  # 切換到該iframe
#             try:
#                 pdf_links_elements = WebDriverWait(driver, 60).until(
#                     EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
#                 )  # 在iframe中等待並查找所有包含".pdf"連結的元素
#                 pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])  # 將找到的連結添加到pdf_links清單中
#             except TimeoutException:
#                 pass  # 如果超時未找到，則忽略
#             driver.switch_to.default_content()  # 切回主文檔內容

#         file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")  # 查找所有包含".pdf"連結的元素
#         for link in file_links_elements:
#             href = link.get_attribute('href')  # 獲取連結的href屬性
#             link_text = link.text.strip()  # 獲取連結的文字並去除首尾空格
#             if href:
#                 pdf_links.append((href, link_text))  # 將連結和文字添加到pdf_links清單中
        
#     except WebDriverException as e:
#         queue.put(f"Error accessing {url} with Selenium: {e}\n")  # 如果發生WebDriver異常，將錯誤信息放入隊列中
#     finally:
#         if driver:
#             driver.quit()  # 關閉WebDriver

#     return pdf_links  # 返回PDF連結清單

def fetch_pdf_links_realtek(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_ti(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_stmicro(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_adi(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_renesas(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_onsemi(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_vishay(url, queue):
    pdf_links = []
    driver = None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(10)
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        while iframes:
            iframe = iframes.pop()
            driver.switch_to.frame(iframe)
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass
            driver.switch_to.default_content()

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        for link in file_links_elements:
            href = link.get_attribute('href')
            link_text = link.text.strip()
            if href:
                pdf_links.append((href, link_text))
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return pdf_links

def fetch_pdf_links_celestica(url, queue):
    # 抓不到url
    """从网站的特定URL中提取PDF文件链接"""
    pdf_links = []  # 用于存储PDF链接的列表

    try:
        # 发送 GET 请求获取网页内容，设置超时时间为10秒
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # 检查请求是否成功

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(response.content, 'html.parser')

        # 找到所有包含 PDF 下载链接的 <a> 标签
        for link in soup.find_all('a', href=True):
            href = link['href']
            # 检查 <a> 标签的 type 属性是否包含 'application/pdf'
            if 'type' in link.attrs and 'pdf' in link['type']:
                full_pdf_link = urljoin("https://corporate.celestica.com", href)
                link_text = link.get_text(strip=True)
                pdf_links.append((full_pdf_link, link_text))
                queue.put(f"Found PDF link: {full_pdf_link} with text: {link_text}\n")  # 将找到的PDF链接放入队列中

    except requests.RequestException as e:
        queue.put(f"Error accessing URL {url}: {e}\n")  # 如果访问URL时发生异常，将错误信息放入队列中

    return pdf_links

def fetch_pdf_links_acer(url):
    """从Acer网站的特定URL中提取PDF文件链接"""
    pdf_links = []  # 用于存储PDF链接的列表
    queue = Queue()  # 创建队列对象

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    def get_response_with_retries(url, retries=3, delay=5):
        for i in range(retries):
            try:
                response = requests.get(url, headers=headers, timeout=20)
                response.raise_for_status()
                return response
            except requests.Timeout:
                queue.put(f"Timeout error accessing URL {url}: The request took too long to complete. Retry {i+1}/{retries}\n")
                time.sleep(delay)
            except requests.RequestException as e:
                queue.put(f"Error accessing URL {url}: {e}\n")
                return None
        return None

    response = get_response_with_retries(url)
    if not response:
        queue.put(f"Failed to access URL {url} after multiple retries.\n")
        while not queue.empty():
            print(queue.get())
        return pdf_links

    # 使用 BeautifulSoup 解析 HTML
    soup = BeautifulSoup(response.content, 'html.parser')

    # 找到特定的 <section> 标签
    sections = soup.find_all('section', class_='section section--padded white')
    if not sections:
        queue.put("Page took too long to load or PDF links not found")
        while not queue.empty():
            print(queue.get())
        return pdf_links
    
    for section in sections:
        # 找到包含PDF链接的 <a> 标签
        link = section.find('a', class_='btn btn--tertiary', href=True)
        if link and 'acer' in link.attrs:
            href = link['href']
            full_pdf_link = href
            link_text = link.get_text(strip=True)
            pdf_links.append((full_pdf_link, link_text))
            queue.put(f"Found PDF link: {full_pdf_link} with text: {link_text}\n")  # 将找到的PDF链接放入队列中

    if not pdf_links:
        queue.put("Page took too long to load or PDF links not found")

    # 打印队列中的信息
    while not queue.empty():
        print(queue.get())

    return pdf_links

def fetch_pdf_links_broadcom(url):
    pdf_links = []  # 存储PDF链接的列表
    retries = 2  # 设置重试次数
    timeout = 30  # 设置超时时间

    try:
        # 发送 GET 请求获取网页内容
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()  # 检查请求是否成功

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(response.content, 'html.parser')

        # 查找包含 "2023" 字段的链接
        links_with_2023 = []
        for link in soup.find_all('a', href=True):
            if '2023' in link.get_text(strip=True):
                full_link = urljoin(url, link['href'])
                links_with_2023.append(full_link)

        # 打印找到的包含 "2023" 字段的链接
        print("Links with '2023' found on the page:")
        for link in links_with_2023:
            print(link)

        # 遍历所有包含 "2023" 字段的链接
        for link in links_with_2023:
            try:
                print(f"Attempting to access URL: {link}")
                # 发送 GET 请求获取链接页面内容
                response_sub = requests.get(link, timeout=timeout)
                response_sub.raise_for_status()  # 检查请求是否成功

                # 使用 BeautifulSoup 解析链接页面内容
                soup_sub = BeautifulSoup(response_sub.content, 'html.parser')

                # 找到所有包含 PDF 下载链接的 <a> 标签
                sub_links = soup_sub.find_all('a', href=True)
                if sub_links is not None:
                    for sub_link in sub_links:
                        if 'type' in sub_link.attrs and sub_link['type'] == 'application/pdf':
                            href = sub_link['href']
                            full_pdf_link = urljoin("https://investors.broadcom.com/", href)
                            pdf_links.append(full_pdf_link)
                            print(f"Found PDF link: {full_pdf_link}")
                else:
                    print("No links found in the sub-page.")

            except requests.RequestException as e:
                print(f"Error accessing URL {link}: {e}")

    except requests.RequestException as e:
        print(f"Error accessing URL {url}: {e}")

    # 打印找到的所有PDF链接
    print("All found PDF links:")
    for link in pdf_links:
        print(link)

    # 尝试下载所有找到的PDF文件
    for link in pdf_links:
        success = False
        for attempt in range(1, retries + 1):
            try:
                print(f"Attempting to download {link}, attempt {attempt} of {retries}")
                response = requests.get(link, timeout=timeout)
                response.raise_for_status()  # 检查请求是否成功
                success = True
                break
            except requests.RequestException as e:
                print(f"Failed to download {link}: {e}, attempt {attempt} of {retries}")

        if not success:
            print(f"Failed to download {link} after {retries} attempts")

    return pdf_links
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #


def fetch_pdf_links_with_dropdown(url, target_year, queue, company_name, downloaded_files_box, root, save_directory):

    """從具有年份下拉菜單的網站中提取PDF文件連結"""
    pdf_links = []  # 用於存儲PDF連結的清單
    driver = None  # 初始化WebDriver為None

    def convert_to_minguo_year(year):
        """將西元年轉換為民國年"""
        return str(int(year) - 1911) + "年"  # 將西元年減去1911得到民國年，並加上"年"字

    minguo_year = convert_to_minguo_year(target_year)  # 將目標年份轉換為民國年

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)  # 使用Firefox啟動WebDriver
        driver.get(url)  # 打開指定的URL

        queue.put(f"Current URL: {driver.current_url}\n")  # 將當前URL放入隊列中
        queue.put(f"Page title: {driver.title}\n")  # 將頁面標題放入隊列中

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 滾動到頁面底部
        time.sleep(5)  # 等待5秒以確保所有元素都加載完成

        try:
            year_dropdown = None  # 初始化年份下拉菜單為None
            try:
                # 嘗試通過ID查找年份下拉菜單
                year_dropdown = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//select[@id='year-select']"))
                )  # 等待並查找ID為'year-select'的下拉菜單元素
                queue.put(f"Found year dropdown by ID\n")  # 找到年份下拉菜單並放入隊列中
            except TimeoutException:
                # 如果未找到，嘗試其他方法
                queue.put(f"Year dropdown not found by ID, trying alternative methods\n")  # 如果未找到，嘗試其他方法

            if not year_dropdown:
                try:
                    # 嘗試通過類別查找年份下拉菜單
                    year_dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'year')]"))
                    )  # 等待並查找類別包含'year'的下拉菜單元素
                    queue.put(f"Found year dropdown by class\n")  # 找到年份下拉菜單並放入隊列中
                except TimeoutException:
                    queue.put(f"Year dropdown not found by class, trying generic select tag\n")  # 如果未找到，嘗試使用通用的<select>標籤

            if not year_dropdown:
                try:
                    # 嘗試通過通用的<select>標籤查找年份下拉菜單
                    year_dropdown = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//select"))
                    )  # 等待並查找通用的<select>標籤元素
                    queue.put(f"Found year dropdown by generic select tag\n")  # 找到年份下拉菜單並放入隊列中
                except TimeoutException:
                    # 如果未找到，列出所有<select>元素
                    queue.put(f"Generic select tag not found, listing all select elements\n")  # 如果未找到，列出所有<select>元素

            if not year_dropdown:
                # 查找所有<select>標籤元素
                select_elements = driver.find_elements(By.TAG_NAME, 'select')  # 查找所有<select>標籤元素
                for select_element in select_elements:
                    # 將找到的<select>元素的外部HTML放入隊列中
                    queue.put(f"Select element found: {select_element.get_attribute('outerHTML')}\n")  # 將找到的<select>元素的外部HTML放入隊列中

            if year_dropdown:
                # 滾動到年份下拉菜單
                driver.execute_script("arguments[0].scrollIntoView(true);", year_dropdown)  # 滾動到年份下拉菜單
                time.sleep(2)  # 等待2秒
                driver.execute_script("arguments[0].style.display = 'block';", year_dropdown)  # 顯示下拉菜單
                
                # 查找所有<option>元素
                options = year_dropdown.find_elements(By.TAG_NAME, 'option')  # 查找所有<option>元素
                for option in options:
                    # 將每個選項的文字放入隊列中
                    queue.put(f"Option: {option.text}\n")  # 將每個選項的文字放入隊列中

                target_option_value = None  # 初始化目標年份選項的值為None
                for option in options:
                    # 查找目標年份的選項
                    if option.text == target_year or option.text == minguo_year:
                        target_option_value = option.get_attribute('value')  # 獲取目標年份選項的值
                        break  # 找到目標選項後退出循環
                
                if target_option_value:
                    try:
                        # 使用JavaScript設置下拉選單的值並觸發變更事件
                        driver.execute_script(f"arguments[0].value = '{target_option_value}'; arguments[0].dispatchEvent(new Event('change'));", year_dropdown)
                        queue.put(f"Set year option via JavaScript: {target_option_value}\n")  # 將設置的選項放入隊列中
                        time.sleep(5)  # 等待5秒以加載新內容

                        queue.put("New content loaded successfully\n")  # 將加載成功的信息放入隊列中

                        # 滾動頁面並嘗試點擊所有可見的PDF下載按鈕和連結
                        scroll_height = driver.execute_script("return document.body.scrollHeight")  # 獲取頁面高度
                        for _ in range(10):  # 假設最多滾動10次
                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 滾動到頁面底部
                            time.sleep(3)  # 等待3秒
                            # 查找所有包含.pdf的連結或包含下載文字的連結
                            pdf_buttons = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')] | //a[contains(text(), '下載')]")
                            queue.put(f"Found {len(pdf_buttons)} potential PDF buttons\n")  # 將找到的PDF按鈕數量放入隊列中
                            for button in pdf_buttons:
                                href = button.get_attribute('href')  # 獲取連結的href屬性
                                if not href:
                                    button.click()  # 如果href屬性不存在，則點擊按鈕
                                    time.sleep(3)  # 等待3秒
                                    # 查找所有包含.pdf或包含下載的連結
                                    pdf_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]")
                                    for link in pdf_links_elements:
                                        href = link.get_attribute('href')  # 獲取連結的href屬性
                                        if href and href not in [link for link, _ in pdf_links]:
                                            pdf_links.append((href, link.text.strip()))  # 將找到的PDF連結添加到清單中
                                            queue.put(f"Found PDF link after click: {href}\n")  # 將找到的PDF連結放入隊列中
                                else:
                                    if href not in [link for link, _ in pdf_links]:
                                        pdf_links.append((href, button.text.strip()))  # 將找到的PDF連結添加到清單中
                                        queue.put(f"Found PDF link: {href}\n")  # 將找到的PDF連結放入隊列中
                            if driver.execute_script("return document.body.scrollHeight") == scroll_height:
                                break  # 如果頁面高度沒有變化，則停止滾動
                            scroll_height = driver.execute_script("return document.body.scrollHeight")  # 更新頁面高度

                        # 確保找到所有PDF連結
                        queue.put(f"Found PDF links: {pdf_links}\n")  # 將找到的PDF連結放入隊列中

                    except WebDriverException as e:
                        queue.put(f"Error clicking year option: {e}\n")  # 如果點擊年份選項時發生異常，將錯誤信息放入隊列中

                else:
                    queue.put(f"Year option {target_year}/{minguo_year} not found in dropdown\n")  # 如果未找到目標年份選項，將信息放入隊列中
            else:
                queue.put(f"Year dropdown or target year {target_year}/{minguo_year} not found on page: {url}\n")  # 如果未找到年份下拉菜單，將信息放入隊列中

        except TimeoutException:
            queue.put(f"Warning: Year dropdown or target year {target_year}/{minguo_year} not found on page: {url}\n")  # 如果超時未找到，將警告信息放入隊列中

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")  # 如果訪問URL時發生異常，將錯誤信息放入隊列中

    finally:
        if driver:
            driver.quit()  # 關閉WebDriver

    # 新增檢查年份的邏輯
    def contains_year(link, year):
        """檢查連結中是否包含指定年份"""
        return bool(re.search(year, link))

    downloaded = False  # 初始化下載狀態為False
    if pdf_links:
        # 確保每個連結都被單獨處理
        for link, text in pdf_links:
            queue.put(f"Processing link: {link}, text: {text}\n")  # 將正在處理的連結和文字放入隊列中
            # 如果 link 是元組則解包
            if isinstance(link, tuple):
                link, text = link
            # 確保連結不包含多個URL
            split_links = link.split(',')
            for split_link in split_links:
                split_link = split_link.strip()
                queue.put(f"Found split link: {split_link}\n")  # 將分割的連結放入隊列中

                # 檢查連結中是否包含年份
                if contains_year(split_link, target_year) or contains_year(split_link, minguo_year):
                    queue.put(f"Link contains the target year: {split_link}\n")
                elif not re.search(r'\d{4}', split_link):  # 檢查連結中是否有四位數字的年份
                    queue.put(f"Link does not contain any year, assuming it is valid: {split_link}\n")
                else:
                    queue.put(f"Link contains a different year, skipping download: {split_link}\n")
                    continue  # 跳過這個連結

                queue.put(f"Attempting to download: {split_link}\n")  # 將嘗試下載的連結放入隊列中
                # 生成公司目錄
                # 生成公司目錄
                company_dir = os.path.join(save_directory, company_name)  # 使用传递的 save_directory

                if not os.path.exists(company_dir):
                    os.makedirs(company_dir)  # 如果目錄不存在，則創建目錄
                # 生成文件名稱，包含公司名稱和安全的文件名稱
                file_name = f"{company_name}_{sanitize_filename(split_link.split('/')[-1])}"  # 生成文件名稱
                # 設定保存路徑
                save_path = os.path.join(company_dir, file_name)  # 設定保存路徑
                if download_file(split_link, save_path):  # 嘗試下載文件
                    downloaded_files_box.insert(END, f"下載完成: {file_name}\n")  # 將下載完成信息插入文本框
                    downloaded_files_box.see(END)  # 滾動到文本框末尾
                    root.update_idletasks()  # 更新界面
                    downloaded = True  # 只要成功下載一個文件，就設置為True
                    queue.put(f"Successfully downloaded: {file_name}\n")  # 將成功下載的信息放入隊列中
                else:
                    queue.put(f"Failed to download: {split_link}\n")  # 如果下載失敗，將錯誤信息放入隊列中

    return downloaded  # 返回下載狀態


def fetch_pdf_links_general(url, target_year, queue, company_name, downloaded_files_box, root):
    """通用函數:從給定URL中提取所有PDF文件連結，並包括靜態文件和HTML文件的檢測"""
    pdf_links = []  # 用於存儲PDF連結的清單
    driver = None  # 初始化WebDriver為None

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)  # 使用Firefox啟動WebDriver
        driver.get(url)  # 打開指定的URL
        
        queue.put(f"Current URL: {driver.current_url}\n")  # 將當前URL放入隊列中作為調試信息
        queue.put(f"Page title: {driver.title}\n")  # 將頁面標題放入隊列中作為調試信息
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 滾動到頁面底部
        time.sleep(5)  # 等待5秒以確保所有元素都加載完成
        
        try:
            # 等待並查找所有包含.pdf或包含下載文字的連結
            pdf_links_elements = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
            )
            # 獲取每個連結的href屬性和文字
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")  # 如果超時未找到PDF連結，將警告信息放入隊列中
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')  # 查找所有iframe元素
        while iframes:
            iframe = iframes.pop()  # 取出最後一個iframe
            driver.switch_to.frame(iframe)  # 切換到該iframe
            try:
                # 在iframe中等待並查找所有包含.pdf或包含下載文字的連結
                pdf_links_elements = WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download')]"))
                )
                # 獲取每個連結的href屬性和文字並添加到pdf_links清單中
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])
            except TimeoutException:
                pass  # 如果超時未找到，則忽略
            iframes.extend(driver.find_elements(By.TAG_NAME, 'iframe'))  # 繼續查找新的iframe元素
            driver.switch_to.default_content()  # 切回主文檔內容

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")  # 如果訪問URL時發生異常，將錯誤信息放入隊列中

    finally:
        if driver:
            driver.quit()  # 關閉WebDriver

    # 應用篩選和下載邏輯
    downloaded = filter_and_download_pdfs(pdf_links, target_year, target_keywords, exclude_keywords, company_name, queue, downloaded_files_box, root)  # 調用篩選和下載PDF的函數
    
    return pdf_links if downloaded else []  # 如果下載成功，返回pdf_links清單，否則返回空清單


def manual_navigate_general(url, queue, target_year, company_name, downloaded_files_box, root):
    """手動導航通用函數: 從給定URL中手動導航並提取所有PDF文件連結"""
    pdf_links = []  # 用於存儲PDF連結的清單
    driver = None  # 初始化WebDriver為None

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)  # 使用Firefox啟動WebDriver
        driver.get(url)  # 打開指定的URL

        time.sleep(15)  # 等待15秒以確保所有元素都加載完成

        keywords = ["investor", "financial", "finance", "reports", "財務", "投資者"]  # 定義關鍵字清單
        for keyword in keywords:
            try:
                queue.put(f"Trying to navigate using keyword: {keyword}\n")  # 將嘗試導航的關鍵字放入隊列中
                # 等待並查找包含關鍵字文字的<a>標籤
                link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//a[contains(text(), '{keyword}')]"))
                )
                link.click()  # 點擊找到的連結
                time.sleep(5)  # 等待5秒以確保頁面加載完成
            except TimeoutException:
                continue  # 如果超時未找到，繼續嘗試下一個關鍵字

        try:
            # 等待並查找所有包含.pdf、下載文字、.html、.aspx、.jsp的連結
            pdf_links_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, '.html') or contains(@href, '.aspx') or contains(@href, '.jsp')]"))
            )
            # 獲取每個連結的href屬性和文字並添加到pdf_links清單中
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]
        except TimeoutException:
            queue.put("Warning: PDF links not found during manual navigation\n")  # 如果超時未找到PDF連結，將警告信息放入隊列中

    except WebDriverException as e:
        queue.put(f"Error during manual navigation: {e}\n")  # 如果手動導航時發生異常，將錯誤信息放入隊列中

    finally:
        if driver:
            driver.quit()  # 關閉WebDriver

    # 應用篩選和下載邏輯
    downloaded = filter_and_download_pdfs(pdf_links, target_year, target_keywords, exclude_keywords, company_name, queue, downloaded_files_box, root)  # 調用篩選和下載PDF的函數
    
    return pdf_links if downloaded else []  # 如果下載成功，返回pdf_links清單，否則返回空清單


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
def decode_url(url):
    """解碼URL中的百分比編碼部分"""
    return urllib.parse.unquote(url)
    # 使用urllib.parse.unquote函數將URL中的百分比編碼部分解碼成普通字符串

def sanitize_filename(filename):
    """去除文件名中的非法字符"""
    return re.sub(r'[\/:*?"<>|]', '_', filename)
    # 使用正則表達式替換文件名中的非法字符（包括 / : * ? " < > |）為下劃線

def download_file(url, save_path, retries=2, timeout=30):
    """下載文件並保存到指定路徑，帶有重試機制"""
    for attempt in range(retries):
        # 設置重試次數
        try:
            response = requests.get(url, headers=headers, allow_redirects=True, timeout=timeout, verify=False)
            # 發送HTTP GET請求以獲取指定URL的內容，允許重定向，設置超時時間
            response.raise_for_status()
            # 如果響應狀態碼表示錯誤，則引發HTTPError異常
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    f.write(response.content)
                    # 將響應內容寫入指定路徑的文件中
                return True
                # 如果成功下載文件，返回True
        except requests.exceptions.RequestException as e:
            print(f"Failed to download {url}: {e}, attempt {attempt + 1} of {retries}")
            # 捕獲HTTP請求異常，並打印錯誤信息
            time.sleep(2)  # 等待2秒后重試
            # 等待2秒后再次嘗試下載
    return False
    # 如果在指定次數的嘗試後仍然無法下載文件，返回False
    
def download_and_save_file(url, link_text, company_name):
    """下載並保存文件，並在文件名前加上公司名稱"""
    company_directory = os.path.join(save_directory, company_name)
    # 構建公司的子文件夾路徑
    if not os.path.exists(company_directory):
        os.makedirs(company_directory)
        # 如果文件夾不存在，則創建它

    decoded_url = decode_url(url)
    # 解碼URL中的百分比編碼部分
    file_name = f"{company_name}_{sanitize_filename(decoded_url.split('/')[-1])}"
    # 使用公司名稱和文件名組合構建新的文件名
    save_path = os.path.join(company_directory, file_name)
    # 構建文件的保存路徑
    print(f"Downloading file from: {decoded_url}")
    # 打印下載文件的URL

    # 增加重試機制
    for attempt in range(3):
        # 設置重試次數為3次
        if download_file(url, save_path):
            # 嘗試下載文件
            print(f"Downloaded: {file_name} to {save_path}")
            # 打印下載成功信息
            return True
            # 返回True表示下載成功
        else:
            print(f"Failed to download: {file_name}, attempt {attempt + 1} of 3")
            # 打印下載失敗信息
            time.sleep(random.uniform(1, 3))
            # 添加1到3秒的隨機延遲後再次嘗試下載
    return False
    # 返回False表示下載失敗

def download_and_save_pdf(url, company_name):
    """下載並保存PDF文件，並在文件名前加上公司名稱"""
    # 創建公司命名的子文件夾
    company_directory = os.path.join(save_directory, company_name)
    # 構建公司的子文件夾路徑
    if not os.path.exists(company_directory):
        os.makedirs(company_directory)
        # 如果文件夾不存在，則創建它

    decoded_url = decode_url(url)
    # 解碼URL中的百分比編碼部分
    file_name = f"{company_name}_{sanitize_filename(decoded_url.split('/')[-1])}"
    # 使用公司名稱和文件名組合構建新的文件名
    save_path = os.path.join(company_directory, file_name)
    # 構建文件的保存路徑
    print(f"Downloading file from: {decoded_url}")
    # 打印下載文件的URL
    if download_file(decoded_url, save_path):
        # 嘗試下載文件
        print(f"Downloaded PDF: {file_name} to {save_path}")
        # 打印下載成功信息
    else:
        print(f"Failed to download: {file_name}")
        # 打印下載失敗信息

def filter_and_download_pdfs(pdf_links, year, keywords, exclude_keywords, company_name, queue, downloaded_files_box, root):
    """根據目標年份和關鍵字篩選並下載PDF文件"""
    pdf_links_filtered = []
    # 初始化一個空列表，用於儲存經過篩選的PDF鏈接
    downloaded = False
    # 標誌是否成功下載至少一個文件

    queue.put(f"Target year: {year}\n")
    # 將目標年份放入佇列，用於日誌記錄或後續處理
    decoded_links = [(decode_url(file_url), link_text) for file_url, link_text in pdf_links]
    # 將PDF鏈接中的百分比編碼部分解碼
    for decoded_url, link_text in decoded_links:
        queue.put(f"Decoded link: {decoded_url} with text: {link_text}\n")
        # 將解碼後的鏈接和文本放入佇列，用於日誌記錄或後續處理

    for decoded_url, link_text in decoded_links:
        queue.put(f"Checking link: {decoded_url} with text: {link_text}\n")
        # 檢查每個解碼後的鏈接和文本
        if any(ex_keyword.lower() in link_text.lower() for ex_keyword in exclude_keywords) or any(ex_keyword.lower() in decoded_url.lower() for ex_keyword in exclude_keywords):
            queue.put(f"Excluded link: {decoded_url}\n")
            # 如果鏈接或文本中包含排除關鍵字，則將其放入佇列，並跳過該鏈接
            continue  # 跳過包含排除關鍵字的鏈接
        year_match = match_year_in_text(link_text, year) or match_year_in_text(decoded_url, year)
        # 檢查鏈接或文本中是否包含目標年份
        keyword_match = match_keywords_in_text(link_text, keywords) or match_keywords_in_text(decoded_url, keywords)
        # 檢查鏈接或文本中是否包含目標關鍵字
        if year_match and keyword_match:
            queue.put(f"Downloading matching link: {decoded_url} with text: {link_text}\n")
            # 如果鏈接或文本同時包含目標年份和關鍵字，則將其放入佇列，並嘗試下載該鏈接
            company_directory = os.path.join(save_directory, company_name)
            # 構建公司的子文件夾路徑
            if not os.path.exists(company_directory):
                os.makedirs(company_directory)
                # 如果文件夾不存在，則創建它
            save_path = os.path.join(company_directory, f"{company_name}_{sanitize_filename(decoded_url.split('/')[-1])}")
            # 構建文件的保存路徑
            if download_file(decoded_url, save_path):
                # 嘗試下載文件
                downloaded_files_box.insert(END, f"下載完成: {company_name}_{sanitize_filename(decoded_url.split('/')[-1])}\n")
                # 將下載完成的信息插入到下載文件框中
                downloaded_files_box.see(END)
                # 滾動到下載文件框的末尾
                root.update_idletasks()
                # 更新根窗口的待處理任務
                downloaded = True
                # 標誌為成功下載
        else:
            pdf_links_filtered.append((decoded_url, link_text))
            # 如果不符合年份或關鍵字，則將其添加到篩選後的PDF鏈接列表中

    for decoded_url, link_text in pdf_links_filtered:
        queue.put(f"Checking filtered link: {decoded_url} with text: {link_text}\n")
        # 檢查篩選後的每個鏈接和文本
        year_match = match_year_in_text(link_text, year) or match_year_in_text(decoded_url, year)
        # 檢查鏈接或文本中是否包含目標年份
        if year_match:
            queue.put(f"Downloading matching filtered link: {decoded_url} with text: {link_text}\n")
            # 如果鏈接或文本包含目標年份，則將其放入佇列，並嘗試下載該鏈接
            company_directory = os.path.join(save_directory, company_name)
            # 構建公司的子文件夾路徑
            if not os.path.exists(company_directory):
                os.makedirs(company_directory)
                # 如果文件夾不存在，則創建它
            save_path = os.path.join(company_directory, f"{company_name}_{sanitize_filename(decoded_url.split('/')[-1])}")
            # 構建文件的保存路徑
            if download_file(decoded_url, save_path):
                # 嘗試下載文件
                downloaded_files_box.insert(END, f"下載完成: {company_name}_{sanitize_filename(decoded_url.split('/')[-1])}\n")
                # 將下載完成的信息插入到下載文件框中
                downloaded_files_box.see(END)
                # 滾動到下載文件框的末尾
                root.update_idletasks()
                # 更新根窗口的待處理任務
                downloaded = True
                # 標誌為成功下載

    return downloaded
    # 返回是否成功下載標誌

def is_pdf(url):
    """檢查文件是否是PDF格式"""
    if url.startswith('javascript'):
        return False  # 跳過javascript鏈接
    try:
        response = requests.head(url, headers=headers, allow_redirects=True, timeout=10, verify=False)
        # 發送HTTP HEAD請求以檢查URL的內容類型，允許重定向，設置超時時間
        content_type = response.headers.get('Content-Type')
        # 獲取響應的Content-Type頭信息
        if content_type is None:
            return False  # 跳過沒有Content-Type頭的鏈接
        return 'pdf' in content_type.lower()
        # 檢查Content-Type頭信息是否包含'pdf'字符串
    except requests.RequestException as e:
        print(f"Error checking if URL is a PDF: {e}", url)
        # 捕獲HTTP請求異常，並打印錯誤信息
        return False

def filter_non_pdf_links(pdf_links):
    """過濾掉非PDF文件的鏈接"""
    filtered_links = []
    # 初始化一個空列表，用於儲存過濾後的PDF鏈接
    for link in pdf_links:
        # 遍歷所有PDF鏈接
        if isinstance(link, tuple) and len(link) == 2:
            # 檢查鏈接是否是元組且長度為2
            url, text = link
            # 解壓鏈接和文本
            if is_pdf(url):
                # 檢查鏈接是否指向PDF文件
                filtered_links.append((url, text))
                # 如果是PDF文件，將其添加到過濾後的鏈接列表中
        else:
            print(f"Unexpected link format: {link}")
            # 如果鏈接格式不符合預期，打印錯誤信息
    return filtered_links
    # 返回過濾後的PDF鏈接列表

def download_all_pdfs(pdf_links, company_name):
    """下載所有檢測到的PDF文件"""
    downloaded = False
    # 標誌是否成功下載至少一個文件
    for file_url in pdf_links:
        # 遍歷所有PDF鏈接
        if download_and_save_pdf(file_url, company_name):
            # 嘗試下載並保存PDF文件
            downloaded = True
            # 如果成功下載至少一個文件，設置標誌為True
    return downloaded
    # 返回是否成功下載標誌

def download_all_files(file_links, company_name):
    """下載所有給定鏈接的文件，並保存到指定目錄"""
    for link in file_links:
        # 遍歷所有文件鏈接
        if isinstance(link, tuple) and len(link) == 2:
            # 檢查鏈接是否是元組且長度為2
            file_url, link_text = link
            # 解壓鏈接和文本
            download_and_save_file(file_url, link_text, company_name)
            # 嘗試下載並保存文件
        else:
            print(f"Skipping invalid link: {link}")
            # 如果鏈接格式不符合預期，打印錯誤信息並跳過該鏈接

class PDFDownloaderApp:
    def __init__(self, root):
        # 初始化函數
        bold_font = font.Font(family="Helvetica", size=10, weight="bold")
        # 設置字體樣式
        
        self.root = root
        # 設置根窗口
        self.root.title("財報擷取系統")
        self.style = ttk.Style()
        
        self.style.theme_use("clam")  # 設置主題樣式
        # 設置樣式為cosmo

        ttk.Label(root, text="输入欲新增的公司:", font=bold_font).pack(pady=5)
        # 創建標籤並設置文字和字體
        self.url_label = ttk.Entry(root, width=80)
        # 創建輸入框並設置寬度
        self.url_label.pack(pady=5)
        # 添加輸入框到窗口

        self.main_frame = ttk.Frame(root)
        # 創建主框架
        self.main_frame.pack()
        # 添加主框架到窗口

        self.output_box = ScrolledText(self.main_frame, wrap='word', height=31, width=82)
        # 創建滾動文本框，設置自動換行、高度和寬度
        self.output_box.grid(row=0, column=0, rowspan=3, padx=10, pady=10)
        # 將滾動文本框添加到主框架的網格佈局中

        self.side_frame = ttk.Frame(self.main_frame)
        # 創建側框架
        self.side_frame.grid(row=0, column=1, padx=10, pady=10)
        # 將側框架添加到主框架的網格佈局中

        self.current_company_box = ScrolledText(self.side_frame, wrap='word', height=4, width=82)
        # 創建滾動文本框，設置自動換行、高度和寬度
        self.current_company_box.pack(pady=5)
        # 添加滾動文本框到側框架
        self.current_company_box.insert(END, "當前正在處理的公司名稱:\n")
        # 在滾動文本框中插入初始文字

        self.failed_company_box = ScrolledText(self.side_frame, wrap='word', height=9, width=82)
        # 創建滾動文本框，設置自動換行、高度和寬度
        self.failed_company_box.pack(pady=5)
        # 添加滾動文本框到側框架
        self.failed_company_box.insert(END, "處理失敗的公司:\n")
        # 在滾動文本框中插入初始文字

        self.downloaded_files_box = ScrolledText(self.side_frame, wrap='word', height=15, width=82)
        # 創建滾動文本框，設置自動換行、高度和寬度
        self.downloaded_files_box.pack(pady=5)
        # 添加滾動文本框到側框架
        self.downloaded_files_box.insert(END, "已經下載的檔案名稱:\n")
        # 在滾動文本框中插入初始文字

        self.button_frame = ttk.Frame(root)
        # 創建按鈕框架
        self.button_frame.pack(pady=10)
        # 添加按鈕框架到窗口

        self.run_button = ttk.Button(self.button_frame, text="Run", command=self.run, bootstyle=SUCCESS)
        # 創建運行按鈕，設置文字和命令
        self.run_button.pack(side='left', padx=5)
        # 添加運行按鈕到按鈕框架的左側

        self.pause_button = ttk.Button(self.button_frame, text="Pause", command=self.pause, bootstyle=WARNING)
        # 創建暫停按鈕，設置文字和命令
        self.pause_button.pack(side='left', padx=5)
        # 添加暫停按鈕到按鈕框架的左側

        self.stop_button = ttk.Button(self.button_frame, text="Stop", command=self.stop, bootstyle=DANGER)
        # 創建停止按鈕，設置文字和命令
        self.stop_button.pack(side='left', padx=5)
        # 添加停止按鈕到按鈕框架的左側

        self.save_button = ttk.Button(self.button_frame, text="Save Output", command=self.save_output, bootstyle=PRIMARY)
        # 創建保存按鈕，設置文字和命令
        self.save_button.pack(side='left', padx=5)
        # 添加保存按鈕到按鈕框架的左側

        self.developer_button = ttk.Button(self.button_frame, text="Developer Options", command=show_developer_window)
        self.developer_button.pack(side='left', padx=5)
        
        self.queue = Queue()
        # 創建佇列
        self.lock = threading.Lock()
        # 創建線程鎖
        self.paused = threading.Event()
        # 創建暫停事件
        self.paused.set()  # Initially not paused
        # 初始化為未暫停
        self.running = threading.Event()
        # 創建運行事件
        self.running.set()  # Initially running
        # 初始化為運行
        self.root.after(100, self.process_queue)
        # 設置在100毫秒後調用process_queue方法
        
        

    def pause(self):
        # 暫停方法
        if self.paused.is_set():
            self.paused.clear()
            # 清除暫停事件
            self.pause_button.config(text="Resume")
            self.queue.put("Process pause\n")
            # 設置暫停按鈕的文字為Resume
        else:
            self.paused.set()
            # 設置暫停事件
            self.pause_button.config(text="Pause")
            self.queue.put("Process resume\n")
            # 設置暫停按鈕的文字為Pause

    def stop(self):
        self.running.clear()
        self.queue.put("Stopping the process...\n")
        self.join_threads()
        self.queue.put("Process stopped\n")
    
    def join_threads(self):
        active_threads = [t for t in threading.enumerate() if t.name.startswith("ThreadPoolExecutor")]
        for t in active_threads:
            t.join(timeout=1)

    def process_site(self, site_name, site_url, target_year, target_keywords, queue, current_company_box, downloaded_files_box, failed_company_box, root, lock, paused):
        try:
            for _ in range(5):  # 假设每个网站需要5步处理
                if not self.running.is_set():
                    return site_name, False
                paused.wait()
                time.sleep(1)

            return site_name, True

        except Exception as e:
            queue.put(f"Error processing {site_name}: {e}\n")
            return site_name, False

    def save_output(self):
        # 保存輸出方法
        output_content = self.output_box.get(1.0, END)
        # 獲取輸出框中的內容
        save_path = os.path.join("C:\PY-CH", "output_log.txt")
        # 設置保存路徑
        with open(save_path, "w", encoding="utf-8") as file:
            file.write(output_content)
            # 將輸出內容寫入文件
        self.queue.put(f"Output has been saved to {save_path}\n")
        # 將保存信息放入佇列

    def download_pdfs(self, new_url):
        # 下載PDF文件方法
        if new_url:
            site_urls['User Added'] = new_url
            # 如果有新URL，則添加到site_urls字典中

        not_downloaded_companies = []
        # 初始化一個列表，用於儲存未下載文件的公司

        companies = list(site_urls.items())
        # 將site_urls字典轉換為列表
        in_progress = []
        # 初始化一個列表，用於儲存正在處理的公司

        with ThreadPoolExecutor(max_workers=1) as executor:
            # 使用ThreadPoolExecutor確保每次只啟動一個實例
            futures = {}
            for i in range(1):  # 每次只處理一個公司
                if companies:
                    site_name, site_url = companies.pop(0)
                    # 彈出一個公司名稱和URL
                    future = executor.submit(process_site, site_name, site_url, target_year, target_keywords, self.queue, self.current_company_box, self.downloaded_files_box, self.failed_company_box, self.root, self.lock, self.paused)
                    # 提交一個任務到線程池中
                    futures[future] = site_name
                    # 將未來對象和公司名稱存儲在字典中
                    in_progress.append(site_name)
                    # 將公司名稱添加到正在處理的列表中
                    self.update_current_companies(in_progress)
                    # 更新正在處理的公司列表

            while futures and self.running.is_set():
                # 當有未來對象且運行標誌設置時
                self.paused.wait()
                # 如果暫停，則等待
                done, _ = as_completed(futures), len(futures)
                # 獲取完成的未來對象
                for future in done:
                    site_name = futures.pop(future)
                    # 從字典中彈出未來對象
                    in_progress.remove(site_name)
                    # 從正在處理的列表中移除公司名稱
                    self.update_current_companies(in_progress)
                    # 更新正在處理的公司列表
                    try:
                        _, downloaded = future.result()
                        # 獲取未來對象的結果
                        # if not downloaded:
                        #     not_downloaded_companies.append(site_name)
                        #     # 如果未下載，則添加到未下載公司列表中
                        #     self.update_failed_companies(site_name)
                        #     # 更新失敗的公司列表
                    except Exception as e:
                        self.queue.put(f"Error processing {site_name}: {e}\n")
                        # 捕獲異常，並將錯誤信息放入佇列
                        # not_downloaded_companies.append(site_name)
                        # # 添加到未下載公司列表中
                        self.update_failed_companies(site_name)
                        # 更新失敗的公司列表

                    if companies:
                        site_name, site_url = companies.pop(0)
                        # 彈出一個公司名稱和URL
                        future = executor.submit(process_site, site_name, site_url, target_year, target_keywords, self.queue, self.current_company_box, self.downloaded_files_box, self.failed_company_box, self.root, self.lock, self.paused)
                        # 提交一個任務到線程池中
                        futures[future] = site_name
                        # 將未來對象和公司名稱存儲在字典中
                        in_progress.append(site_name)
                        # 將公司名稱添加到正在處理的列表中
                        self.update_current_companies(in_progress)
                        # 更新正在處理的公司列表

        # self.queue.put("The following companies had no downloads:\n")
        # # 將未下載公司信息放入佇列
        # self.queue.put(f"{not_downloaded_companies}\n")
        # # 將未下載公司列表放入佇列
    
    def update_current_companies(self, in_progress):
        # 更新當前正在處理的公司列表
        self.root.after(0, self._update_current_companies, in_progress)
        # 在主線程中調用更新方法

    def _update_current_companies(self, in_progress):
        # 實際更新當前正在處理的公司列表的方法
        self.current_company_box.delete(1.0, END)
        # 清空當前公司文本框
        self.current_company_box.insert(END, "當前正在處理的公司名稱:\n")
        # 插入初始文字
        for company in in_progress:
            self.current_company_box.insert(END, f"{company}\n")
            # 插入每個正在處理的公司名稱
        self.current_company_box.see(END)
        # 滾動到文本框的末尾
        self.root.update_idletasks()
        # 更新根窗口的待處理任務

    def update_failed_companies(self, site_name):
        # 更新失敗的公司列表
        self.root.after(0, self._update_failed_companies, site_name)
        # 在主線程中調用更新方法

    def _update_failed_companies(self, site_name):
        # 實際更新失敗的公司列表的方法
        self.failed_company_box.insert(END, f"{site_name}\n")
        # 插入失敗的公司名稱
        self.failed_company_box.see(END)
        # 滾動到文本框的末尾
        self.root.update_idletasks()
        # 更新根窗口的待處理任務

    def run(self):
        # 運行方法
        new_url = self.url_label.get()
        # 獲取輸入框中的新URL
        self.queue.put("Starting the download process...\n")
        # 將啟動下載過程的信息放入佇列
        self.running.set()
        # 確保運行標誌設置為True
        threading.Thread(target=self.download_pdfs, args=(new_url,)).start()
        # 創建一個新線程來下載PDF文件

    def process_queue(self):
        # 處理佇列的方法
        try:
            while True:
                line = self.queue.get_nowait()
                # 獲取佇列中的一行
                self.output_box.insert(END, line)
                # 將該行插入到輸出框
                self.output_box.see(END)
                # 滾動到文本框的末尾
                self.root.update_idletasks()
                # 更新根窗口的待處理任務
        except Empty:
            self.root.after(10, self.process_queue)
            # 如果佇列為空，10毫秒後再次調用process_queue方法
        except Exception as e:
            self.queue.put(f"Error processing queue: {e}\n")    

        

def process_site(site_name, site_url, target_year, target_keywords, queue, current_company_box, downloaded_files_box, failed_company_box, root, lock, paused):
    with lock:  # 使用鎖確保線程安全
        current_company_box.insert(END, f"正在處理: {site_name}\n")  # 在current_company_box中插入當前正在處理的公司名稱
        current_company_box.see(END)  # 滾動至current_company_box的末尾
        root.update_idletasks()  # 更新界面

        downloaded = False  # 初始化下載狀態為False
        
        if site_name == 'NXP':
            download_pdfs_for_year_nxp(target_year, {site_name: site_url}, target_keywords, queue, downloaded_files_box, root)
            return site_name, True  # NXP 下載完成，返回True
        elif site_name == 'AMD':
            queue.put("Calling fetch_pdf_links_amd\n")  
            links = fetch_pdf_links_amd(site_url, target_year, target_keywords, queue)
        elif site_name == 'Qualcomm':
            queue.put("Calling fetch_pdf_links_qualcomm\n")  
            links = fetch_pdf_links_qualcomm(site_url, target_year, target_keywords, queue)
        elif site_name == 'Nanya':
            queue.put("Calling fetch_pdf_links_nanya\n")  
            links = fetch_pdf_links_nanya(site_url, target_year, target_keywords, queue)
        elif site_name == 'Broadcom':
            queue.put("Calling fetch_pdf_links_broadcom\n")  
            links = fetch_pdf_links_broadcom(site_url)
        elif site_name == 'Realtek':
            queue.put("Calling fetch_pdf_links_realtek\n")  
            links = fetch_pdf_links_realtek(site_url, queue)
        elif site_name == 'TI':
            queue.put("Calling fetch_pdf_links_ti\n")  
            links = fetch_pdf_links_ti(site_url, queue)
        elif site_name == 'ST Micro':
            queue.put("Calling fetch_pdf_links_stmicro\n")  
            links = fetch_pdf_links_stmicro(site_url, queue)
        elif site_name == 'ADI':
            queue.put("Calling fetch_pdf_links_adi\n")  
            links = fetch_pdf_links_adi(site_url, queue)
        elif site_name == 'Renesas':
            queue.put("Calling fetch_pdf_links_renesas\n")  
            links = fetch_pdf_links_renesas(site_url, queue)
        elif site_name == 'Onsemi':
            queue.put("Calling fetch_pdf_links_onsemi\n")  
            links = fetch_pdf_links_onsemi(site_url, queue)
        elif site_name == 'Vishay':
            queue.put("Calling fetch_pdf_links_vishay\n")  
            links = fetch_pdf_links_vishay(site_url, queue)
        elif site_name == 'Wistron':
            queue.put("Calling fetch_pdf_links_wistron\n")  
            links = fetch_pdf_links_wistron(site_url, queue)
        elif site_name == 'HonHai':
            queue.put("Calling fetch_pdf_links_honhai\n")  
            links = fetch_pdf_links_honhai(site_url, queue)
            links = [(decode_url(url), text) for url, text in links]  # 解碼連結
        elif site_name == 'Inventec':
            queue.put("Calling fetch_pdf_links_inventec\n")  
            links = fetch_pdf_links_inventec(site_url, queue, target_year, downloaded_files_box, root)
            if links:
                for link, text in links:
                    if download_and_save_pdf(link, text, site_name):  # 下載並保存PDF文件
                        downloaded_files_box.insert(END, f"下載完成: {site_name}_{sanitize_filename(link.split('/')[-1])}\n")
                        downloaded_files_box.see(END)  # 滾動至末尾
                        root.update_idletasks()  # 更新界面
                return site_name, True  # Inventec 下載完成，返回True
            else:
                return site_name, False  # Inventec 下載失敗，返回False
        elif site_name == 'Quanta':
            queue.put("Calling fetch_pdf_links_quanta\n")  
            links = fetch_pdf_links_quanta(site_url, queue)
        elif site_name == 'Pegatron':
            queue.put("Calling fetch_pdf_links_pegatron\n")  
            links = fetch_pdf_links_pegatron(site_url, queue)
        elif site_name == 'Compal':
            queue.put("Calling fetch_pdf_links_compal\n")  
            links = fetch_pdf_links_compal(site_url, queue)
        elif site_name == 'Wiwynn':
            queue.put("Calling fetch_pdf_links_wiwynn\n")  
            links = fetch_pdf_links_wiwynn(site_url, queue)
        elif site_name == 'Liteon':
            queue.put("Calling fetch_pdf_links_liteon\n")  
            links = fetch_pdf_links_liteon(site_url, queue, site_name, save_directory)
            # links = [(decode_url(url), text) for url, text in links]  # 解碼連結
        elif site_name == 'Chicony':
            queue.put("Calling fetch_pdf_links_chicony\n")  
            links = fetch_pdf_links_chicony(queue)
        elif site_name == 'Intel':
            queue.put("Calling fetch_pdf_links_intel\n")  
            links = fetch_pdf_links_intel(site_url, queue)
        elif site_name == 'Celestica':
            queue.put("Calling fetch_pdf_links_celestica\n")  
            links = fetch_pdf_links_celestica(site_url, queue)
        elif site_name == 'Acer':
            queue.put("Calling fetch_pdf_links_acer\n")  
            links = fetch_pdf_links_acer(site_url)
        elif site_name == 'Micron':
            queue.put("Calling fetch_pdf_links_micron\n")  
            links = fetch_pdf_links_micron(site_url, queue)
        else:
            # 首先嘗試使用年份下拉菜單提取PDF連結
            queue.put(f"Calling fetch_pdf_links_with_dropdown for {site_name}\n")  
            downloaded = fetch_pdf_links_with_dropdown(site_url, target_year, queue, site_name, downloaded_files_box, root, save_directory)

            if not downloaded:
                queue.put(f"Dropdown not found or no links found for {site_name}, trying fetch_pdf_links_general\n")
                downloaded = fetch_pdf_links_general(site_url, target_year, queue, site_name, downloaded_files_box, root)
            
            if not downloaded:
                queue.put(f"No links found for {site_name}, trying manual navigation\n")
                # 手動導航到目標頁面
                manual_navigation_successful = manual_navigate_general(site_url, queue, target_year, site_name, downloaded_files_box, root)
                if manual_navigation_successful:
                    # 手動導航成功後，再次嘗試使用年份下拉菜單提取PDF連結
                    queue.put(f"Manual navigation successful, calling fetch_pdf_links_with_dropdown again for {site_name}\n")
                    downloaded = fetch_pdf_links_with_dropdown(site_url, target_year, queue, site_name, downloaded_files_box, root, save_directory)
                    
                    if not downloaded:
                        queue.put(f"Dropdown not found or no links found after manual navigation for {site_name}, trying fetch_pdf_links_general again\n")
                        downloaded = fetch_pdf_links_general(site_url, target_year, queue, site_name, downloaded_files_box, root)

            if not downloaded:
                queue.put(f"No links found for {site_name}\n")
                failed_company_box.insert(END, f"{site_name}\n")  # 將失敗公司名稱插入failed_company_box中
                failed_company_box.see(END)  # 滾動至末尾
                root.update_idletasks()  # 更新界面
                return site_name, False  # 返回False表示未找到連結

        if site_name not in no_filter_companies and not downloaded:
            # 如果公司名稱不在無需篩選的公司列表中
            links = filter_non_pdf_links(links)  # 篩選非PDF連結
            downloaded = filter_and_download_pdfs(links, target_year, target_keywords, exclude_keywords, site_name, queue, downloaded_files_box, root)
            # 篩選並下載符合條件的PDF文件
        elif not downloaded:
            # 如果公司名稱在無需篩選的公司列表中
            downloaded = download_all_pdfs(links, site_name)  # 下載所有PDF文件
        
        if not downloaded:
            failed_company_box.insert(END, f"{site_name}\n")  # 將失敗公司名稱插入failed_company_box中
            failed_company_box.see(END)  # 滾動至末尾
            root.update_idletasks()  # 更新界面
        
        return site_name, downloaded  # 返回公司名稱和下載狀態

        
if __name__ == "__main__":
    # 主程序入口
    root = Tk()
    # 创建Tkinter根窗口
    app = PDFDownloaderApp(root)
    # 创建PDF下载应用程序实例

    
    root.mainloop()
    # 进入Tkinter主循环