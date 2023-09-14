from ChromeDriver.Chromedriver_version import ChromeDriver
import time
driver = ChromeDriver("win64").create_driver()
driver.get("https://google.com")
time.sleep(10)