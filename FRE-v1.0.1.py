#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# test5-v4.py
# @Author :  ()
# @Link   : 
# @Date   : 2024/6/20

# 修復南亞Nanya，改採fetch內置下載邏輯以避開download變量設置問題


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
from selenium.webdriver.firefox.options import Options as FirefoxOptions  # 匯入FirefoxOptions，用於設定Firefox瀏覽器選項
from selenium.webdriver.common.by import By  # 匯入By模組，用於定位網頁元素
from selenium.webdriver.support.ui import WebDriverWait  # 匯入WebDriverWait，用於顯式等待
from selenium.webdriver.support import expected_conditions as EC  # 匯入預期條件模組，用於顯式等待的條件設置
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, WebDriverException # 匯入Selenium的例外處理，用於處理各種異常情況
from urllib.parse import urljoin  
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
from queue import Queue, Empty
from concurrent.futures import ThreadPoolExecutor, as_completed
import webbrowser
from datetime import datetime
import platform
import sys
from selenium import webdriver
from selenium.webdriver.firefox.service import Service

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

# 使用 webdriver_manager 自动下载和配置 geckodriver
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)


# hook-selenium.py
from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('selenium')

# 获取可执行文件所在目录
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# 驱动程序路径
geckodriver_path = os.path.join(base_path, 'geckodriver.exe')

# 使用 Service 类指定驱动程序路径
service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service)


# 修改ThreadPoolExecutor的最大工作數量，將其設置為更小的值
MAX_WORKERS = 3

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 要檢查的網站URL列表
site_urls = {
        
    

    
    
    'NXP': 'https://investors.nxp.com/financial-information/financial-information-0',
    'Qualcomm': 'https://investor.qualcomm.com/financial-information/historical-financial-results',
    'AMD': 'https://ir.amd.com/financial-information/historical-financials',
    'HonHai': 'https://www.foxconn.com/zh-tw/investor-relations/financial-information/reports?category=quarterly',
    'Liteon': 'https://www.liteon.com/zh-tw/investor/financialreports/8',
    'Compal': 'https://www.compal.com/investor-relations/financial-release/#consolidated-financial',
    'Pegatron': 'https://www.pegatroncorp.com/investorRelation/downloadLibrary/type/financial_information/year/all',
    'Wistron': 'https://www.wistron.com/ch/Investors/Financials/FinancialReports',
    
    'Quanta': 'https://www.quantatw.com/Quanta/chinese/investment/financials_qr3.aspx',
    'Wiwynn': 'https://www.wiwynn.com/zh/investors#quarterlyresults',
    'Benchmark': 'https://ir.bench.com/financials/quarterly-results/default.aspx',
    'celestica': 'https://corporate.celestica.com/quarterly-results',
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
    'Acer': 'https://www.acer.com/corporate/zh/investor-relations/financials/annual-reports',
    
    'MPS': 'https://www.monolithicpower.com/en/about-mps/investor-relations/press-releases.html',

    'AOS': 'https://investor.aosmd.com/financial-information/annual-reports-and-proxy/default.aspx',

    'SK hynix': 'https://www.skhynix.com/ir/UI-FR-IR12_T4',
    
    'Inventec': 'https://www.inventec.com/tw/finance-2',
    
    
    'Intel': 'https://www.intc.com/financial-info/financial-results',
    'Nvidia': 'https://investor.nvidia.com/financial-info/financial-reports/default.aspx',
    'Samsung': 'https://www.samsung.com/global/ir/financial-information/audited-financial-statements/',
    
    'Micron': 'https://investors.micron.com/quarterly-results',
    'Kioxia': 'https://www.kioxia-holdings.com/en-jp/about/company.html#anc05',
    'Winbond': 'https://www.winbond.com/hq/about-winbond/investor/financial-information/financial-reports/?__locale=zh_TW',
    'Nanya': 'https://www.nanya.com/tw/IR/39',
    'Broadcom': 'https://investors.broadcom.com/financial-information/quarterly-results',
    'Realtek': 'https://www.realtek.com/InvestorRelations/FinancialStatements?menu_id=461&lang=zh-TW',
    'Mediatek': 'https://corp.mediatek.tw/investor-relations/financial-information/quarterly-earnings/2023',
    'TI': 'https://investor.ti.com/financial-information/earnings-annual-reports',
    'ST Micro': 'https://investors.st.com/financial-information/quarterly-results',
    'Infineon': 'https://www.infineon.com/cms/en/about-infineon/investor/reports-and-presentations/#financial-results',
    'ADI': 'https://investor.analog.com/financial-info/quarterly-results',
    'Renesas': 'https://www.renesas.com/us/en/about/investor-relations/event/presentation/2023#financial_results',
    'Onsemi': 'https://investor.onsemi.com/financials',
    'Toshiba': 'https://www.global.toshiba/ww/ir/corporate/finance/er/er-list.html',
    'Vishay': 'https://ir.vishay.com/financial-information/quarterly-results',

    'Ams-Osram': 'https://ams-osram.com/about-us/investor-relations/financial-results-and-reports',
    'Marvell': 'https://investor.marvell.com/quarterly-results',
    'Skyworks': 'https://investors.skyworksinc.com/annual-reports-and-proxies',
    'Qorvo': 'https://ir.qorvo.com/press-releases',
    'Rohm': 'https://www.rohm.com/ir/library/financial-report/',
    'Diodes': 'https://investor.diodes.com/financial-performance-1',
    
}

# 保存PDF文件的目錄
save_directory = 'C:\\PY-CH\\PY-CH-11'

# 檢查保存目錄是否存在,如果不存在則創建它
if not os.path.exists(save_directory):
    os.makedirs(save_directory)

# 設置目標年份和季度
# target_year = '2023'  # 目標年份
target_keywords = ['Results', 'Release', 'finance-statement', '財務', 'Earnings Release', 'Financial Results', 'Quarterly Results', 'Annual Report', 'Consolicated Financial']  # 目標關鍵字
exclude_keywords = ['逐字稿', '簡報', 'Presentation']  # 排除的關鍵字
no_filter_companies = ['Quanta']  # 不需要過濾的公司

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
webdriver_service = FirefoxService(r"C:\Program Files\Mozilla Firefox\geckodriver-v0.34.0-win64\geckodriver.exe")
# 設定geckodriver的路徑，這是Firefox的WebDriver，負責與Firefox瀏覽器進行交互。

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
        print(f"Error checking if URL is a PDF: {e}")
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

def fetch_pdf_links_liteon(queue):
    # 函數從光寶網站中提取PDF文件鏈接
    pdf_links = []
    # 初始化一個空列表，用於儲存找到的PDF鏈接
    driver = None
    # 初始化driver變數

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        # 啟動Firefox瀏覽器並訪問指定的URL
        driver.get('https://www.liteon.com/zh-tw/investor/financialreports/8')
        # 訪問光寶網站的財務報告頁面
        
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
        pdf_links_elements = pdf_links_elements[:12]
        # 只取前12個鏈接元素
        pdf_links = [(decode_url(link.get_attribute('href')), link.text.strip()) for link in pdf_links_elements]
        # 提取這些鏈接的href屬性和文本，並解碼href，添加到pdf_links列表中
        
        queue.put(f"Liteon found links: {pdf_links}\n")
        # 將找到的PDF鏈接放入佇列，用於日誌記錄或後續處理
        
    except WebDriverException as e:
        queue.put(f"Error accessing Liteon site with Selenium: {e}\n")
        # 處理WebDriver異常，並將錯誤信息放入佇列

    finally:
        if driver:
            driver.quit()
            # 確保在結束時關閉瀏覽器

    return pdf_links
    # 返回找到的PDF鏈接列表

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

def fetch_pdf_links_micron(url, queue):
    driver = None
    pdf_links = set()
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)
        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")

        # 确保页面加载完成
        time.sleep(10)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)

        # 获取页面上所有的链接
        links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
        for link in links:
            href = link.get_attribute('href')
            if href and ('.pdf' in href or 'download' in href):
                pdf_links.add(href)
                queue.put(f"Found PDF link: {href}\n")
        
        # 检查页面中所有的 iframe
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        for iframe in iframes:
            driver.switch_to.frame(iframe)
            links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
            for link in links:
                href = link.get_attribute('href')
                if href and ('.pdf' in href or 'download' in href):
                    pdf_links.add(href)
                    queue.put(f"Found PDF link in iframe: {href}\n")
            driver.switch_to.default_content()

        # 处理动态加载的内容，尝试点击“Load More”按钮或其他动态加载的元素
        while True:
            try:
                load_more_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Load More')]")
                driver.execute_script("arguments[0].click();", load_more_button)
                time.sleep(5)

                # 再次获取页面上所有的链接
                links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
                for link in links:
                    href = link.get_attribute('href')
                    if href and ('.pdf' in href or 'download' in href):
                        pdf_links.add(href)
                        queue.put(f"Found PDF link: {href}\n")
            except:
                break

        # 再次检查页面上的所有链接，确保没有遗漏
        links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf') or contains(@href, 'download') or contains(@href, 'financials')]")
        for link in links:
            href = link.get_attribute('href')
            if href and ('.pdf' in href or 'download' in href):
                pdf_links.add(href)
                queue.put(f"Found PDF link: {href}\n")

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")
    finally:
        if driver:
            driver.quit()

    return list(pdf_links)


def fetch_pdf_links_nanya(url, target_year, queue, downloaded_files_box, root, save_directory):
    pdf_links = []
    driver = None

    def convert_to_minguo_year(year):
        return str(int(year) - 1911) + "年"

    minguo_year = convert_to_minguo_year(target_year)

    def download_and_save_pdf(url, filename, directory):
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()
            if not os.path.exists(directory):
                os.makedirs(directory)
            filepath = os.path.join(directory, filename)
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        except requests.exceptions.RequestException as e:
            queue.put(f"Error downloading {url}: {e}\n")
            return False

    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)
        driver.get(url)

        queue.put(f"Current URL: {driver.current_url}\n")
        queue.put(f"Page title: {driver.title}\n")

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)

        try:
            year_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='dropdownMenu1']"))
            )
            year_button.click()

            queue.put("Year dropdown opened successfully\n")

            def get_target_option():
                options = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//ul[@class='select-option dropdown-menu']/li/a"))
                )
                target_option = None
                for option in options:
                    option_text = option.text.strip()
                    queue.put(f"Option found: {option_text}\n")
                    if option_text and (option_text == target_year or option_text == minguo_year):
                        target_option = option
                return target_option

            retries = 3
            target_option = None
            while retries > 0:
                try:
                    target_option = get_target_option()
                    if target_option:
                        queue.put(f"Target option found: {target_option.text}\n")
                        break
                    else:
                        queue.put(f"Year option {target_year}/{minguo_year} not found\n")
                        return pdf_links
                except StaleElementReferenceException:
                    queue.put("StaleElementReferenceException encountered, retrying...\n")
                    time.sleep(2)
                    retries -= 1

            if target_option:
                retries = 3
                while retries > 0:
                    try:
                        year_button.click()
                        time.sleep(1)
                        
                        options = WebDriverWait(driver, 10).until(
                            EC.presence_of_all_elements_located((By.XPATH, "//ul[@class='select-option dropdown-menu']/li/a"))
                        )
                        
                        for option in options:
                            option_text = option.text.strip()
                            if option_text == target_year or option_text == minguo_year:
                                option.click()
                                queue.put(f"Clicked year option: {option_text}\n")
                                break

                        WebDriverWait(driver, 10).until(
                            EC.url_contains(f"{target_year}")
                        )

                        time.sleep(5)
                        break
                    except (StaleElementReferenceException, TimeoutException):
                        queue.put("Exception encountered while clicking or waiting for URL change, retrying...\n")
                        retries -= 1
                        time.sleep(2)

                if retries == 0:
                    queue.put("Failed to click year option after several retries\n")
                    return pdf_links

                body_text = driver.find_element(By.TAG_NAME, 'body').text
                if target_year not in body_text and minguo_year not in body_text:
                    queue.put(f"Error: Did not navigate to the correct year page. Content does not match target year.\n")
                    return pdf_links

                scroll_height = driver.execute_script("return document.body.scrollHeight")
                for _ in range(10):
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(3)
                    try:
                        download_wrap = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "download-wrap"))
                        )
                        pdf_buttons = download_wrap.find_elements(By.XPATH, ".//a[contains(@href, 'Get_IRFinancialReport_FileName')]")
                        queue.put(f"Found {len(pdf_buttons)} PDF elements\n")
                        for button in pdf_buttons:
                            href = button.get_attribute('href')
                            text = button.text.strip()
                            full_url = urljoin(url, href)
                            if href and href not in [link for link, _ in pdf_links]:
                                filename = f"{sanitize_filename(text)}.pdf"
                                if download_and_save_pdf(full_url, filename, save_directory):
                                    downloaded_files_box.insert(END, f"下載完成: {filename}\n")
                                    downloaded_files_box.see(END)
                                    root.update_idletasks()
                                    pdf_links.append((full_url, text))
                                    queue.put(f"Found and downloaded PDF link: {full_url} with text: {text}\n")
                        if driver.execute_script("return document.body.scrollHeight") == scroll_height:
                            break
                        scroll_height = driver.execute_script("return document.body.scrollHeight")
                    except StaleElementReferenceException as e:
                        queue.put(f"Error during PDF button search: {e}\n")
                        continue

                queue.put(f"Found PDF links: {pdf_links}\n")

        except TimeoutException:
            queue.put(f"Warning: Year dropdown or target year {target_year}/{minguo_year} not found on page: {url}\n")

    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")

    finally:
        if driver:
            driver.quit()

    return pdf_links



def fetch_pdf_links_broadcom(url, queue):
    pdf_links = []  # 用於存儲PDF連結的清單
    driver = None  # 初始化WebDriver為None
    try:
        driver = webdriver.Firefox(service=webdriver_service, options=firefox_options)  # 使用Firefox啟動WebDriver
        driver.get(url)  # 打開指定的URL
        
        queue.put(f"Current URL: {driver.current_url}\n")  # 將當前URL放入隊列中
        queue.put(f"Page title: {driver.title}\n")  # 將頁面標題放入隊列中
        
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 滾動到頁面底部
        time.sleep(10)  # 等待10秒以確保所有元素都加載完成
        
        try:
            pdf_links_elements = WebDriverWait(driver, 60).until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
            )  # 等待並查找所有包含".pdf"連結的元素
            pdf_links = [(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements]  # 獲取每個連結的href屬性和文字
        except TimeoutException:
            queue.put(f"Warning: PDF links not found on page: {url}\n")  # 如果超時未找到PDF連結，將警告信息放入隊列中
        
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')  # 查找所有iframe元素
        while iframes:
            iframe = iframes.pop()  # 取出最後一個iframe
            driver.switch_to.frame(iframe)  # 切換到該iframe
            try:
                pdf_links_elements = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )  # 在iframe中等待並查找所有包含".pdf"連結的元素
                pdf_links.extend([(link.get_attribute('href'), link.text.strip()) for link in pdf_links_elements])  # 將找到的連結添加到pdf_links清單中
            except TimeoutException:
                pass  # 如果超時未找到，則忽略
            driver.switch_to.default_content()  # 切回主文檔內容

        file_links_elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")  # 查找所有包含".pdf"連結的元素
        for link in file_links_elements:
            href = link.get_attribute('href')  # 獲取連結的href屬性
            link_text = link.text.strip()  # 獲取連結的文字並去除首尾空格
            if href:
                pdf_links.append((href, link_text))  # 將連結和文字添加到pdf_links清單中
        
    except WebDriverException as e:
        queue.put(f"Error accessing {url} with Selenium: {e}\n")  # 如果發生WebDriver異常，將錯誤信息放入隊列中
    finally:
        if driver:
            driver.quit()  # 關閉WebDriver

    return pdf_links  # 返回PDF連結清單


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


import os
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import time

def fetch_pdf_links_ti(url, target_year, queue, downloaded_files_box, root, save_directory):
    pdf_links = []
    driver = None

    try:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36')

        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'})
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        queue.put("Initialized Chrome WebDriver with anti-detection measures\n")
        
        driver.get(url)
        queue.put(f"Opened URL: {url}\n")

        # 等待頁面加載
        time.sleep(10)

        # 滾動頁面直到找到目標年份的部分
        target_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, f"//a[contains(text(), '{target_year} Earnings')]"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(target_element).perform()
        queue.put(f"Scrolled to {target_year} Earnings section\n")

        # 點擊展開目標年份的部分
        target_element.click()
        queue.put(f"Clicked on {target_year} Earnings section\n")

        # 等待展開的內容加載
        time.sleep(5)

        # 查找所有季度
        quarters = ['Q1', 'Q2', 'Q3', 'Q4']
        for quarter in quarters:
            queue.put(f"Processing {quarter} {target_year}\n")
            
            try:
                quarter_section = driver.find_element(By.XPATH, f"//h4[contains(text(), '{quarter}')]")
                pdf_elements = quarter_section.find_elements(By.XPATH, "./following-sibling::div//a[text()='PDF']")
                for element in pdf_elements:
                    href = element.get_attribute('href')
                    parent_div = element.find_element(By.XPATH, "./ancestor::div[contains(@class, 'QRsItems')]")
                    category = parent_div.find_element(By.XPATH, ".//div[contains(@class, 'field-content')]").text
                    if href:
                        file_name = f"TI_{target_year}_{quarter}_{category.replace(' ', '_')}.pdf"
                        pdf_links.append((href, file_name))
                        queue.put(f"Found PDF link for {quarter} {category}: {href}\n")
            except NoSuchElementException:
                queue.put(f"Could not find section for {quarter}\n")

        # 查找年度報告
        try:
            annual_report = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//div[contains(@id, '{target_year}-annual-report')]//a[contains(@class, 'aRep')]"))
            )
            href = annual_report.get_attribute('href')
            if href:
                file_name = f"TI_{target_year}_Annual_Report.pdf"
                pdf_links.append((href, file_name))
                queue.put(f"Found annual report link: {href}\n")
        except (TimeoutException, NoSuchElementException) as e:
            queue.put(f"Error finding annual report link: {str(e)}\n")

        # # 下載PDF文件
        # for link, file_name in pdf_links:
        #     try:
        #         response = requests.get(link)
        #         if response.status_code == 200:
        #             file_path = os.path.join(save_directory, file_name)
        #             with open(file_path, 'wb') as file:
        #                 file.write(response.content)
        #             queue.put(f"Downloaded: {file_name}\n")
        #             downloaded_files_box.insert("end", f"{file_name}\n")
        #             downloaded_files_box.see("end")
        #             root.update_idletasks()
        #             downloaded = True
        #         else:
        #             queue.put(f"Failed to download {file_name}. Status code: {response.status_code}\n")
        #     except Exception as e:
        #         queue.put(f"Error downloading {file_name}: {str(e)}\n")

    except Exception as e:
        queue.put(f"Unexpected error: {str(e)}\n")

    finally:
        if driver:
            driver.quit()
            queue.put("Closed Chrome WebDriver\n")

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
                target_option = None  # 初始化目標年份選項為None
                for option in options:
                    # 查找目標年份的選項
                    if option.text == target_year or option.text == minguo_year:
                        target_option = option  # 查找目標年份的選項
                        break  # 找到目標選項後退出循環
                if target_option:
                    # 滾動到目標年份選項
                    driver.execute_script("arguments[0].scrollIntoView(true);", target_option)  # 滾動到目標年份選項
                    time.sleep(2)  # 等待2秒
                    driver.execute_script("window.scrollBy(0, -100);")  # 向上滾動一點，以確保不被遮擋

                    try:
                        # 點擊目標年份選項
                        target_option.click()  # 點擊目標年份選項
                        queue.put(f"Clicked year option: {target_option.text}\n")  # 將點擊的選項放入隊列中

                        # 觸發下拉菜單的change事件
                        driver.execute_script("arguments[0].dispatchEvent(new Event('change'))", year_dropdown)  # 觸發下拉菜單的change事件
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

def download_and_save_pdf(url, link_text, company_name):
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
        print(f"Error checking if URL is a PDF: {e}")
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
    for file_url, link_text in pdf_links:
        # 遍歷所有PDF鏈接
        if download_and_save_pdf(file_url, link_text, company_name):
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

import ctypes
from ctypes import wintypes

# # 定义 DWM_BLURBEHIND 结构
# class DWM_BLURBEHIND(ctypes.Structure):
#     _fields_ = [("dwFlags", wintypes.DWORD),
#                 ("fEnable", wintypes.BOOL),
#                 ("hRgnBlur", wintypes.HRGN),
#                 ("fTransitionOnMaximized", wintypes.BOOL)]

#     DWM_BB_ENABLE = 0x00000001


class PDFDownloaderApp:
    
    def __init__(self, root):
        bold_font = font.Font(family="Helvetica", size=10, weight="bold")
        # 初始化 target_year
        self.target_year = '2024'
        
        
        
        self.root = root
        self.root.title("財報擷取系統")
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
                
         # 创建主框架并设置居中对齐
        container_frame = ttk.Frame(root)
        container_frame.pack(pady=10, padx=10, fill='x')

        # 创建左侧、居中、右侧框架
        left_frame = ttk.Frame(container_frame, width=50)
        left_frame.pack(side='left', expand=True)

        center_frame = ttk.Frame(container_frame)
        center_frame.pack(side='left', expand=True)

        right_frame = ttk.Frame(container_frame, width=100)
        right_frame.pack(side='right', expand=True)

        # 添加"输入欲新增的公司"标签和输入框到居中框架
        ttk.Label(center_frame, text="输入欲新增的公司:", font=bold_font).pack(side='left', pady=5)
        self.url_label = ttk.Entry(center_frame, width=50)
        self.url_label.pack(side='left', pady=5, padx=5)

        # 添加"擷取年份"标签和下拉选单到右侧框架
        ttk.Label(right_frame, text="擷取年份:", font=bold_font).pack(side='left', pady=7)
        self.year_combobox = ttk.Combobox(right_frame, values=[str(year) for year in range(2010, 2031)], width=10)
        self.year_combobox.set('2024')
        self.year_combobox.pack(side='left', pady=5)
        self.year_combobox.bind("<<ComboboxSelected>>", self.update_target_year)
        

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack()

        self.output_box = ScrolledText(self.main_frame, wrap='word', height=31, width=82)
        self.output_box.grid(row=0, column=0, rowspan=3, padx=10, pady=10)

        self.side_frame = ttk.Frame(self.main_frame)
        self.side_frame.grid(row=0, column=1, padx=10, pady=10)

        self.current_company_box = ScrolledText(self.side_frame, wrap='word', height=4, width=82)
        self.current_company_box.pack(pady=5)
        self.current_company_box.insert(END, "當前正在處理的公司名稱:\n")

        self.failed_company_box = ScrolledText(self.side_frame, wrap='word', height=9, width=82)
        self.failed_company_box.pack(pady=5)
        self.failed_company_box.insert(END, "處理失敗的公司:\n")

        self.downloaded_files_box = ScrolledText(self.side_frame, wrap='word', height=15, width=82)
        self.downloaded_files_box.pack(pady=5)
        self.downloaded_files_box.insert(END, "已經下載的檔案名稱:\n")

        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10)

        self.run_button = ttk.Button(self.button_frame, text="Run", command=self.run, bootstyle=SUCCESS)
        self.run_button.pack(side='left', padx=5)

        self.pause_button = ttk.Button(self.button_frame, text="Pause", command=self.pause, bootstyle=WARNING)
        self.pause_button.pack(side='left', padx=5)

        self.stop_button = ttk.Button(self.button_frame, text="Stop", command=self.stop, bootstyle=DANGER)
        self.stop_button.pack(side='left', padx=5)

        self.save_button = ttk.Button(self.button_frame, text="Save Output", command=self.save_output, bootstyle=PRIMARY)
        self.save_button.pack(side='left', padx=5)

        self.developer_button = ttk.Button(self.button_frame, text="Developer Options", command=show_developer_window)
        self.developer_button.pack(side='left', padx=5)
        
        self.feedback_button = ttk.Button(self.button_frame, text="意見反映", command=self.send_feedback)
        self.feedback_button.pack(side='left', padx=5)
        
        # 檢查最新版本
        latest_version, download_url = check_github_version()
        current_version = "v1.0.0"  # 您的目前版本號

        # 顯示版本提示
        if latest_version and latest_version != current_version:
            version_label = tk.Label(self.button_frame, text=f"新版本可用: {latest_version}", fg="#FF0000", font=("Arial", 9, "bold"), cursor="hand2")
            version_label.pack(side='right', padx=10)
            version_label.bind("<Button-1>", lambda e: self.download_new_version(download_url))
        else:
            version_label = tk.Label(self.button_frame, text=f"當前版本: {current_version}", fg="black", font=("Arial", 8))
            version_label.pack(side='right', padx=10)
        
        # 设置初始模式为浅色模式
        self.is_dark_mode = False
        self.apply_style()

        
        # 根据模式设置版本提示标签的样式
        if self.is_dark_mode:
            version_label.configure(bg='#2e2e2e', fg='#ffffff')
        else:
            version_label.configure(bg='#ffffff', fg='#000000')

        self.queue = Queue()
        self.lock = threading.Lock()
        self.paused = threading.Event()
        self.paused.set()  # Initially not paused
        self.running = threading.Event()
        self.running.set()  # Initially running
        self.root.after(100, self.process_queue)

        # 在 button_frame 中添加切换模式按钮
        self.toggle_mode_button = ttk.Button(self.button_frame, text="切换深淺色", command=self.toggle_mode, style="TButton")
        self.toggle_mode_button.pack(side='left', padx=5)

        
    
    def toggle_mode(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_style()



    def apply_style(self):

        if self.is_dark_mode:
            self.root.configure(bg='#2e2e2e')
            self.style.configure('TFrame', background='#2e2e2e')
            self.style.configure('TLabel', background='#2e2e2e', foreground='#ffffff')
            self.style.configure('TButton', background='#5e5e5e', foreground='#ffffff')
            self.style.configure('TEntry', fieldbackground='#3e3e3e', foreground='#ffffff')
            self.style.configure('TText', background='#3e3e3e', foreground='#ffffff')
            self.output_box.configure(background='#3e3e3e', foreground='#ffffff')
            self.current_company_box.configure(background='#3e3e3e', foreground='#ffffff')
            self.failed_company_box.configure(background='#3e3e3e', foreground='#ffffff')
            self.downloaded_files_box.configure(background='#3e3e3e', foreground='#ffffff')
            
            # 设置版本号区域的样式
            self.style.configure('Version.TLabel', background='#2e2e2e', foreground='#ffffff')
            for widget in self.button_frame.winfo_children():
                if isinstance(widget, tk.Label):
                    widget.configure(bg='#2e2e2e', fg='#ffffff')
        else:
            self.root.configure(bg='#ffffff')  # 设置浅色模式的白色背景
            self.style.configure('TFrame', background='#ffffff')  # 使用白色背景
            self.style.configure('TLabel', background='#ffffff', foreground='#000000')
            self.style.configure('TButton', background='#d0d0d0', foreground='#000000')
            self.style.configure('TEntry', fieldbackground='#ffffff', foreground='#000000')
            self.style.configure('TText', background='#ffffff', foreground='#000000')
            self.output_box.configure(background='#ffffff', foreground='#000000')
            self.current_company_box.configure(background='#ffffff', foreground='#000000')
            self.failed_company_box.configure(background='#ffffff', foreground='#000000')
            self.downloaded_files_box.configure(background='#ffffff', foreground='#000000')
            
            # 设置版本号区域的样式
            self.style.configure('Version.TLabel', background='#ffffff', foreground='#000000')
            for widget in self.button_frame.winfo_children():
                if isinstance(widget, tk.Label):
                    widget.configure(bg='#ffffff', fg='#000000')
            


        
        
                
    def update_target_year(self, event):
        self.target_year = self.year_combobox.get()



    def download_new_version(self, url):
        import webbrowser
        webbrowser.open_new(url)

    def send_feedback(self):
        current_version = "v1.0.0"  # 您的目前版本號
        today_date = datetime.today().strftime('%Y-%m-%d')
        subject = f"版本 {current_version} 意見反映 - {today_date}"
        body = "請在此輸入您的意見..."
        gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to=rio91527@gmail.com&su={subject}&body={body}"
        webbrowser.open_new(gmail_url)   
        
    def pause(self):
        # 暫停方法
        if self.paused.is_set():
            self.paused.clear()
            # 清除暫停事件
            self.pause_button.config(text="Resume")
            # 設置暫停按鈕的文字為Resume
        else:
            self.paused.set()
            # 設置暫停事件
            self.pause_button.config(text="Pause")
            # 設置暫停按鈕的文字為Pause

    def stop(self):
        # 停止方法
        self.running.clear()
        # 清除運行事件
        self.queue.put("Stopping the process...\n")
        # 將停止信息放入佇列

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
                    future = executor.submit(process_site, site_name, site_url, self.target_year, target_keywords, self.queue, self.current_company_box, self.downloaded_files_box, self.failed_company_box, self.root, self.lock, self.paused)                    # 提交一個任務到線程池中
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
                        future = executor.submit(process_site, site_name, site_url, self.target_year, target_keywords, self.queue, self.current_company_box, self.downloaded_files_box, self.failed_company_box, self.root, self.lock, self.paused)                        # 提交一個任務到線程池中
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
        # 在运行前读取输入框中的值并设置目标年份
        target_year = self.year_combobox.get() or '2024'  # 默认为 2024
        print(f"目标年份设为: {self.target_year}")
        
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
            links = fetch_pdf_links_nanya(site_url, target_year, queue, downloaded_files_box, root, save_directory)
            queue.put(f"Fetched links for Nanya: {links}\n")
            downloaded = True if links else False
        elif site_name == 'Micron':
            queue.put("Calling fetch_pdf_links_micron\n")  
            links = fetch_pdf_links_micron(site_url, queue)
        elif site_name == 'Broadcom':
            queue.put("Calling fetch_pdf_links_broadcom\n")  
            links = fetch_pdf_links_broadcom(site_url, queue)
        elif site_name == 'Realtek':
            queue.put("Calling fetch_pdf_links_realtek\n")  
            links = fetch_pdf_links_realtek(site_url, queue)
        elif site_name == 'TI':
            links = fetch_pdf_links_ti(site_url, target_year, queue, downloaded_files_box, root, save_directory)
            queue.put(f"Fetched links for TI: {links}\n")
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
            links = fetch_pdf_links_liteon(queue)
            links = [(decode_url(url), text) for url, text in links]  # 解碼連結
        elif site_name == 'Chicony':
            queue.put("Calling fetch_pdf_links_chicony\n")  
            links = fetch_pdf_links_chicony(queue)
        
        elif site_name == 'Acer':
            queue.put("Calling fetch_pdf_links_acer\n")  
            links = fetch_pdf_links_acer(queue)
        elif site_name == 'Intel':
            queue.put("Calling fetch_pdf_links_intel\n")  
            links = fetch_pdf_links_intel(site_url, queue)
            
                       
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

    
# 檢查GitHub最新版本
def check_github_version():
    url = "https://api.github.com/repos/rio10255254/Financial-Report-Extraction-App/releases/latest"
    try:
        response = requests.get(url)
        response.raise_for_status()
        latest_release = response.json()
        latest_version = latest_release["tag_name"]
        download_url = latest_release["assets"][0]["browser_download_url"] if latest_release["assets"] else latest_release["html_url"]
        return latest_version, download_url
    except requests.RequestException as e:
        print(f"Error checking latest release: {e}")
        return None, None



if __name__ == "__main__":
    # 主程序入口
    root = Tk()
    # 创建Tkinter根窗口
    app = PDFDownloaderApp(root)
    # 创建PDF下载应用程序实例
   
    
    root.mainloop()
    # 进入Tkinter主循环