from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import os, time

class WebDriver:
    def create(platformName, deviceName):
        match platformName.lower():
            case "google": 
                from ChromeDriver.Chromedriver_version import ChromeDriver
                return ChromeDriver(deviceName.lower()).create_driver()
            case "edge": 
                from EdgeDriver.Edgedriver_version import EdgeDriver
                return EdgeDriver(deviceName.lower()).create_driver()
    def initialize():
        import shutil
        shutil.rmtree("./TEMP/")