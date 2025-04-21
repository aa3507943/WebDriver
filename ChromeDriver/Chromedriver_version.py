import requests
from bs4 import BeautifulSoup
import platform
import platform, subprocess
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import os, time, json

if platform.system() == 'Windows':
    from win32com.client import Dispatch
    import pythoncom  

class ChromeDriverForMac:
    def __init__(self, system):
        self.originalPath = os.getcwd()
        currentPath = os.path.abspath(__file__)
        currentDir = os.path.dirname(currentPath)
        os.chdir(currentDir)
        with open("url.json") as js:
            self.data = json.load(js)
            print("Data", self.data)
        self.System = system
        self.download_chromedriver()
        self.extract_zip()

    def climb_chromedriver_version(self):
        import json
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        verList = []
        curVersion = self.get_chrome_version()
        heads = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'}
        proxy={'http': 'http://10.10.10.10:8000', 'https': 'http://10.10.10.10:1212'}
        response = requests.get(self.data['chromedriver-resourse'], headers= heads)
        with open("./TEMP/download_link.json", mode= "w") as f:
            f.write(response.text)
        with open("./TEMP/download_link.json", mode= "r") as f:
            links = json.load(f)
            for i in range(len(links['versions'])):
                if curVersion in links['versions'][i]['version']:
                    verList.append(links['versions'][i]['version'])
            
        print("爬到的ChromeDriver版本為", verList[-1])
        print(f"{self.data['chromedriver-url']}/{verList[-1]}/{self.System}/chromedriver-{self.System}.zip")
        return f"{self.data['chromedriver-url']}/{verList[-1]}/{self.System}/chromedriver-{self.System}.zip"

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
        os.chdir(f"./TEMP/chromedriver-{self.System}")
        os.system("xattr -d com.apple.quarantine chromedriver")

    def get_chrome_version(self):
        if platform.system() == 'Darwin':  # 检查是否是Mac OS
            cmd = "/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version"
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output, error = process.communicate()
            version = output.decode('utf-8').strip()
            version = version[14:20]
            print("當前機器版本為: ", version)
            return version
        else:
            return "Chrome version retrieval is only supported on Mac OS"
    
    def create_driver(self):
        try:
            options = webdriver.ChromeOptions()
            prefs = {
            'profile.default_content_setting_values':
            {
                'notifications': 2,
            },
            'profile.default_content_settings.popups': 0, 
            "profile.default_content_setting_values.clipboard": 1,
            'download.default_directory': os.path.abspath(''),
            "download.prompt_for_download": False,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False
            }
            options.add_experimental_option('prefs', prefs)
            # options.add_argument('--disable-gpu')
            options.add_argument('--lang=zh-TW')
            # options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--window-size=800x600')
            options.add_argument('--mute-audio')
            # options.add_argument('--start-minimized')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--remote-debugging-port=9222')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--disable-features=InterestCohort')
            options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
            # options.add_argument('--log-level=1')
            #抓出與本機端chrome相同版本的版號
            if self.System in ["win32", "win64"]:
                driverName = "chromedriver.exe"
            else:
                driverName = "chromedriver"
            path = os.path.abspath(driverName)
            print(path)
            os.system(f"chmod +x {path}")
            # driver = webdriver.Chrome(service= Service(executable_path= f"./TEMP/chromedriver-{self.System}/{driverName}"), options= options)
            driver = webdriver.Chrome(service= Service(executable_path= path, log_output="chromedriver.log"),  options= options)
            # driver = webdriver.Chrome(options= options)
            # ExceptionHandler(msg= "Successfully open browser driver. 成功開啟瀏覽器驅動器", exceptionLevel= "info")
            os.chdir(self.originalPath)
            return driver
        except:
            # ExceptionHandler(msg= "Cannot open browser driver. 無法開啟瀏覽器驅動器", exceptionLevel= "critical")
            pass

class ChromeDriverForWindows:
    def __init__(self, system):
        self.originalPath = os.getcwd()
        currentPath = os.path.abspath(__file__)
        currentDir = os.path.dirname(currentPath)
        os.chdir(currentDir)
        with open("url.json") as js:
            self.data = json.load(js)
            print("Data", self.data)
        self.System = system
        self.download_chromedriver()
        self.extract_zip()


    def climb_chromedriver_version(self):
        import json
        if not os.path.isdir("./TEMP"):
            os.makedirs("./TEMP")
        verList = []
        curVersion = self.get_chrome_version()
        response = requests.get(self.data['chromedriver-resourse'])
        with open("./TEMP/download_link.json", mode= "w") as f:
            f.write(response.text)
        with open("./TEMP/download_link.json", mode= "r") as f:
            links = json.load(f)
            for i in range(len(links['versions'])):
                if curVersion in links['versions'][i]['version']:
                    verList.append(links['versions'][i]['version'])
        print("爬到的ChromeDriver版本為", verList[-1])
        print(f"{self.data['chromedriver-url']}/{verList[-1]}/{self.System}/chromedriver-{self.System}.zip")
        return f"{self.data['chromedriver-url']}/{verList[-1]}/{self.System}/chromedriver-{self.System}.zip"

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
  
        pythoncom.CoInitialize()
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        print("電腦當前Chrome版本為", version)
        return version

    def get_chrome_version(self): #將機器的chrome版本最後一位去除
        self.chromeVersion = list(filter(None, [self.get_version_via_com(p) for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]]))[0]
        tmp = self.chromeVersion.split(".")
        for i in range(2):
            tmp.pop()
        self.chromeVersion = ".".join(tmp)
        return self.chromeVersion + "."
    
    def create_driver(self):
        try:
            options = webdriver.ChromeOptions()
            prefs = {
            'profile.default_content_setting_values':
            {
                'notifications': 2,
            },
            'profile.default_content_settings.popups': 0, 
            "profile.default_content_setting_values.clipboard": 1,
            'download.default_directory': os.path.abspath(''),
            "download.prompt_for_download": False,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False
            }
            options.add_experimental_option('prefs', prefs)
            # options.add_argument('--disable-gpu')
            options.add_argument('--lang=zh-TW')
            # options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--window-size=800x600')
            options.add_argument('--mute-audio')
            # options.add_argument('--start-minimized')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--remote-debugging-port=9222')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--disable-features=InterestCohort')
            options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
            options.add_argument('--log-level=1')
            #抓出與本機端chrome相同版本的版號
            if self.System in ["win32", "win64"]:
                driverName = "chromedriver.exe"
            else:
                driverName = "chromedriver"
            # path = os.path.abspath(f"./TEMP/chrome/driver-win64/{driverName}")
            # path = os.path(f"./TEMP/chrome/driver-win64/{driverName}")
            path = os.path.join(os.path.dirname(__file__), f"TEMP/chromedriver-win64/{driverName}")
            print(path)
            # driver = webdriver.Chrome(service= Service(executable_path= f"./TEMP/chromedriver-{self.System}/{driverName}", service_args=["--verbose", "--enable-logging", "--log-level=0" ]), options= options)
            driver = webdriver.Chrome(service= Service(executable_path=path, log_output= "chromedriver.log"), options= options)
            # driver = webdriver.Chrome(service= Service(executable_path= path))
            # ExceptionHandler(msg= "Successfully open browser driver. 成功開啟瀏覽器驅動器", exceptionLevel= "info")
            os.chdir(self.originalPath)
            return driver
        except:
            # ExceptionHandler(msg= "Cannot open browser driver. 無法開啟瀏覽器驅動器", exceptionLevel= "critical")
            pass