import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import shutil

# Options for Chrome WebDriver
op = Options()
op.add_argument('--disable-notifications')
op.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
path = os.getcwd()

# create data folder

Created_path = path + '\data'
try:
    os.mkdir(Created_path)
except OSError:
    print("Creation of the directory %s failed" % path)
else:
    print("Successfully created the directory %s " % path)

# Initializing the Chrome webdriver with the options
driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

# Opening the page
driver.get("http://www.hbagahfund.com/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"Agas details.xlsx"))
# Closing the webdriver
driver.close()
