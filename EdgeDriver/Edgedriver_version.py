import requests
from bs4 import BeautifulSoup
from win32com.client import Dispatch
from selenium import webdriver
from selenium.webdriver.edge.service import Service
import xml.etree.ElementTree as ET
import os, time


class EdgeDriver:
    def __init__(self, system):
        self.System = system
        self.download_edgedriver()
        self.extract_zip()

    def __version_filter(self, data):
        startIdx = data.index(" ")
        endIdx = data.index(":")
        return data[startIdx+1:endIdx]
 

    def climb_edgedriver_version(self):
        import json
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        verList = []
        response = requests.get("https://github.com/MicrosoftDocs/Edge-Enterprise/blob/public/edgeenterprise/microsoft-edge-relnote-beta-channel.md")
        with open("./TEMP/download_link.json", mode= "w") as f:
            f.write(response.text)
        with open("./TEMP/download_link.json", mode= "r") as f:
            links = json.load(f)
            for i in range(len(links['payload']['blob']['headerInfo']['toc'])):
                if self.get_edge_version() in links['payload']['blob']['headerInfo']['toc'][i]['text']:
                   verList.append(self.__version_filter(links['payload']['blob']['headerInfo']['toc'][i]['text']))
        return f"https://msedgedriver.azureedge.net/{verList[0]}/edgedriver_{self.System}.zip"
    

    def download_edgedriver(self):
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        response = requests.get(self.climb_edgedriver_version())
        with open (f"./TEMP/edgedriver_{self.System}.zip", 'wb') as file:
            file.write(response.content)
        startTime = time.time()
        while True:
            if os.path.isfile(f"./TEMP/edgedriver_{self.System}.zip"):
                break
            else:
                if time.time() - startTime() > 60:
                    raise
                else:
                    continue

    def extract_zip(self):
        import zipfile
        file = zipfile.ZipFile(f"./TEMP/edgedriver_{self.System}.zip")
        file.extractall(path= f"./TEMP/")
        file.close()
        startTime = time.time()
        while True:
            if os.path.isfile(f"./TEMP/msedgedriver.exe"):
                break
            else:
                if time.time() - startTime > 60:
                    raise
                else:
                    continue

    def get_version_via_com(self, filename): #從目標路徑取得該機器上的Chrome當前版本
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        return version

    def get_edge_version(self): #將機器的edge版本最後一位去除
        self.edgeVersion = list(filter(None, [self.get_version_via_com(p) for p in [r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"]]))[0]
        tmp = self.edgeVersion.split(".")
        for i in range(2):
            tmp.pop()
        self.edgeVersion = ".".join(tmp)
        # print(self.edgeVersion)
        return self.edgeVersion + "."
    
    def create_driver(self):
        try:
            options = webdriver.EdgeOptions()
            prefs = {
            'profile.default_content_setting_values':
            {
                'notifications': 2,
            },
            'profile.default_content_settings.popups': 0, 
            'download.default_directory': os.path.abspath('TEMP'),
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
                driverName = "msedgedriver.exe"
            else:
                driverName = "msedgedriver"
            driver = webdriver.Edge(service= Service(executable_path= f"./TEMP/{driverName}"), options= options)
            # driver = webdriver.Chrome(options= options)
            # ExceptionHandler(msg= "Successfully open browser driver. 成功開啟瀏覽器驅動器", exceptionLevel= "info")
            return driver
        except:
            # ExceptionHandler(msg= "Cannot open browser driver. 無法開啟瀏覽器驅動器", exceptionLevel= "critical")
            pass
