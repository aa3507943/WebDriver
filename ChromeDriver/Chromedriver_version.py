import requests
from bs4 import BeautifulSoup
from win32com.client import Dispatch
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import os, time

# class DownloadFitChromeDrive:
#     def __init__(self):
#         self.get_chrome_version()
#         self.climb_chromedriver_version()
#         self.arrange_version_list()
#         self.downloadVersion = self.get_chromedriver_download_version()
        
#     def climb_chromedriver_version(self):  #從chromedriver下載頁面上爬出所有chromedriver版本
#         response = requests.get("https://chromedriver.chromium.org/downloads")
#         rep = BeautifulSoup(response.text, "html.parser")
#         self.versionList = []
#         tmp = {}
#         # print(rep.select("p[dir='ltr'] a[href][target='_blank'] span"))
#         for i in rep.select("p[dir='ltr'] a[href][target='_blank'] span"):
#             if i.text != "release notes":
#                 self.versionList.append(i.text)
#         tmp = tmp.fromkeys(self.versionList)
#         self.versionList = list(tmp.keys())
#         self.finalList = []
#         for i in range(len(self.versionList)):
#             if "ChromeDriver" in self.versionList[i]:
#                 self.finalList.append(self.versionList[i].replace("ChromeDriver ", ""))
#         # print(self.finalList)

#     def arrange_version_list(self): #整理並歸納爬下來的chromedriver版本，存出兩個串列，一是去掉最後一位的版號，另一是完整版號
#         tmp = ""
#         self.compareVersionList = []
#         for i in range(len(self.finalList)):
#                 tmp = self.finalList[i].split(".")
#                 tmp.pop()
#                 self.compareVersionList.append(".".join(tmp))
#         # print(self.compareVersionList)
 
#     def get_version_via_com(self, filename): #從目標路徑取得該機器上的Chrome當前版本
#         parser = Dispatch("Scripting.FileSystemObject")
#         try:
#             version = parser.GetFileVersion(filename)
#         except Exception:
#             return None
#         return version

#     def get_chrome_version(self): #將機器的chrome版本最後一位去除
#         self.chromeVersion = list(filter(None, [self.get_version_via_com(p) for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
#                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]]))[0]
#         tmp = self.chromeVersion.split(".")
#         tmp.pop()
#         self.chromeVersion = ".".join(tmp)
#         return self.chromeVersion
#     # def get_chrome_version(self):
#     #     tree = ET.parse(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.VisualElementsManifest.xml")
#     #     root = tree.getroot()
#     #     self.chromeVersion = root.find('VisualElements').get('Square150x150Logo').split("\\")[0]
#     #     tmp = self.chromeVersion.split(".")
#     #     tmp.pop()
#     #     self.chromeVersion = ".".join(tmp)
#         # print("chrome version:", self.chromeVersion)
    
#     def get_chromedriver_download_version(self): #把機器上去掉最後一位的版號與chromedriver去掉一位版號的串列作搜尋，抓出對應index，並取得完整版號
#         if self.chromeVersion in self.compareVersionList:
#             downloadVersion = self.finalList[self.compareVersionList.index(self.chromeVersion)]
#             return downloadVersion


# class ChromeDriver:
#     def __init__(self, system):
#         self.System:str = system
#         self.download_chromedriver()
#         self.extract_zip()

#     def climb_chromedriver_version(self):
#         response = requests.get("https://googlechromelabs.github.io/chrome-for-testing/#stable")
#         rep = BeautifulSoup(response.text, "html.parser")
#         parseData = rep.select("div[class='table-wrapper'] td code")
#         for i in parseData:
#             if f"chromedriver-{self.System}" in i.text and self.get_chrome_version() in i.text:
#                 # print(i.text)
#                 return i.text

#     def download_chromedriver(self):
#         if not os.path.isdir("./TEMP"):
#             os.makedirs("./TEMP")
#         response = requests.get(self.climb_chromedriver_version())
#         with open (f"./TEMP/chromedriver.zip", 'wb') as file:
#             file.write(response.content)
#         startTime = time.time()
#         while True:
#             if os.path.isfile(f"./TEMP/chromedriver.zip"):
#                 break
#             else:
#                 if time.time() - startTime > 60:
#                     raise
#                 else:
#                     continue

#     def extract_zip(self):
#         import zipfile
#         file = zipfile.ZipFile(f"./TEMP/chromedriver.zip")
#         file.extractall(path= f"./TEMP/")
#         file.close()
#         startTime = time.time()
#         while True:
#             if os.path.isdir(f"./TEMP/chromedriver-{self.System}"):
#                 break
#             else:
#                 if time.time() - startTime() > 60:
#                     raise
#                 else:
#                     continue

#     def get_version_via_com(self, filename): #從目標路徑取得該機器上的Chrome當前版本
        # import pythoncom
        # pythoncom.CoInitialize()
#         parser = Dispatch("Scripting.FileSystemObject")
#         try:
#             version = parser.GetFileVersion(filename)
#         except Exception:
#             return None
#         return version

#     def get_chrome_version(self): #將機器的chrome版本最後一位去除
#         self.chromeVersion = list(filter(None, [self.get_version_via_com(p) for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
#                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]]))[0]
#         tmp = self.chromeVersion.split(".")
#         for i in range(2):
#             tmp.pop()
#         self.chromeVersion = ".".join(tmp)
#         # print(self.chromeVersion)
#         return self.chromeVersion
    
#     def create_driver(self):
#         try:
#             options = webdriver.ChromeOptions()
#             prefs = {
#             'profile.default_content_setting_values':
#             {
#                 'notifications': 2,
#             },
#             'profile.default_content_settings.popups': 0, 
#             'download.default_directory': os.path.abspath('AzureDevops\\CSV Folder\\'),
#             "download.prompt_for_download": False,
#             "safebrowsing_for_trusted_sources_enabled": False,
#             "safebrowsing.enabled": False}
#             options.add_experimental_option('prefs', prefs)
#             options.add_argument('--disable-gpu')
#             options.add_argument('--lang=zh-TW')
#             # options.add_argument('--headless')
#             options.add_argument('--no-sandbox')
#             options.add_argument('--windows-size=800x600')
#             # options.add_argument('--start-minimized')
#             options.add_argument('--disable-dev-shm-usage')
#             options.add_argument('--remote-debugging-port=9222')
#             options.add_argument('--disable-blink-features=AutomationControlled')
#             options.add_argument('--disable-features=InterestCohort')
#             # options.add_argument('--log-level=1')
#             #抓出與本機端chrome相同版本的版號
#             if self.System in ["win32", "win64"]:
#                 driverName = "chromedriver.exe"
#             else:
#                 driverName = "chromedriver"
#             driver = webdriver.Chrome(service= Service(executable_path= f"./TEMP/chromedriver-{self.System}/{driverName}"), options= options)
#             # driver = webdriver.Chrome(options= options)
#             # ExceptionHandler(msg= "Successfully open browser driver. 成功開啟瀏覽器驅動器", exceptionLevel= "info")
#             return driver
#         except:
#             # ExceptionHandler(msg= "Cannot open browser driver. 無法開啟瀏覽器驅動器", exceptionLevel= "critical")
#             pass



class ChromeDriver:
    def __init__(self, system):
        self.System:str = system
        self.download_chromedriver()
        self.extract_zip()

    def climb_chromedriver_version(self):
        import json
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        verList = []
        response = requests.get("https://raw.githubusercontent.com/GoogleChromeLabs/chrome-for-testing/main/data/known-good-versions-with-downloads.json")
        with open("./TEMP/download_link.json", mode= "w") as f:
            f.write(response.text)
        with open("./TEMP/download_link.json", mode= "r") as f:
            links = json.load(f)
            for i in range(len(links['versions'])):
                if self.get_chrome_version() in links['versions'][i]['version']:
                    verList.append(links['versions'][i]['version'])
        # print(verList[-1])
        return f"https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{verList[-1]}/win64/chromedriver-{self.System}.zip"

    def download_chromedriver(self):
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        response = requests.get(self.climb_chromedriver_version())
        with open (f"./TEMP/chromedriver.zip", 'wb') as file:
            file.write(response.content)
        startTime = time.time()
        while True:
            if os.path.isfile(f"./TEMP/chromedriver.zip"):
                break
            else:
                if time.time() - startTime > 60:
                    raise
                else:
                    continue

    def extract_zip(self):
        import zipfile
        file = zipfile.ZipFile(f"./TEMP/chromedriver.zip")
        file.extractall(path= f"./TEMP/")
        file.close()
        startTime = time.time()
        while True:
            if os.path.isdir(f"./TEMP/chromedriver-{self.System}"):
                break
            else:
                if time.time() - startTime() > 60:
                    raise
                else:
                    continue

    def get_version_via_com(self, filename): #從目標路徑取得該機器上的Chrome當前版本
        import pythoncom
        pythoncom.CoInitialize()
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        return version

    def get_chrome_version(self): #將機器的chrome版本最後一位去除
        self.chromeVersion = list(filter(None, [self.get_version_via_com(p) for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]]))[0]
        tmp = self.chromeVersion.split(".")
        for i in range(2):
            tmp.pop()
        self.chromeVersion = ".".join(tmp)
        # print(self.chromeVersion)
        return self.chromeVersion
    
    def create_driver(self):
        try:
            options = webdriver.ChromeOptions()
            prefs = {
            'profile.default_content_setting_values':
            {
                'notifications': 2,
            },
            'profile.default_content_settings.popups': 0, 
            'download.default_directory': os.path.abspath('AzureDevops\\CSV Folder\\'),
            "download.prompt_for_download": False,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False}
            options.add_experimental_option('prefs', prefs)
            options.add_argument('--disable-gpu')
            options.add_argument('--lang=zh-TW')
            # options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--windows-size=800x600')
            # options.add_argument('--start-minimized')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--remote-debugging-port=9222')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--disable-features=InterestCohort')
            # options.add_argument('--log-level=1')
            #抓出與本機端chrome相同版本的版號
            if self.System in ["win32", "win64"]:
                driverName = "chromedriver.exe"
            else:
                driverName = "chromedriver"
            driver = webdriver.Chrome(service= Service(executable_path= f"./TEMP/chromedriver-{self.System}/{driverName}"), options= options)
            # driver = webdriver.Chrome(options= options)
            # ExceptionHandler(msg= "Successfully open browser driver. 成功開啟瀏覽器驅動器", exceptionLevel= "info")
            return driver
        except:
            # ExceptionHandler(msg= "Cannot open browser driver. 無法開啟瀏覽器驅動器", exceptionLevel= "critical")
            pass
