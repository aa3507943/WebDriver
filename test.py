from ChromeDriver.Chromedriver_version import ChromeDriver
from EdgeDriver.Edgedriver_version import EdgeDriver
import time
# driver = ChromeDriver("win64").create_driver()
driver1 = EdgeDriver("win64").create_driver()
# driver.get("https://google.com")
driver1.get("https://youtube.com")
time.sleep(10)
