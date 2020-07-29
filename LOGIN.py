__author__ = 'tatsuya'
import configparser
import os
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from pandas import json_normalize

options = {
    'log-level':'error'
}
configur = configparser.ConfigParser()
#Fungsi Buka File Config
def write_file():
    configur.write(open('config.ini', 'w'))
#Fungsi Generate File Config
def set_file():
  configur['data'] = {'username': '', 'password': '', 'url': 'alamat.sipd.kemendagri.go.id', 'tahun': '2021', 'status': '501', 'session': '', 'currenturl': ''}

#Cek keberadaan file config
if not os.path.exists('config.ini'):
  set_file()
  write_file()
else:
  configur.read('config.ini')

#Function Login:
def login(driver):
  #Definisi konfigurasi file yang akan diproses
  try:
    username = configur.get('data', 'username')
    password = configur.get('data', 'password')
    url = configur.get('data', 'url')
    #tahun = configur.get('data', 'tahun')
    status = configur.get('data', 'status')
    driver.get("https://{}/daerah/main/0/landing".format(url))
    #print("username: {}".format(username))
    driver.set_window_size(1226, 808)
    WebDriverWait(driver, 3).until(expected_conditions.presence_of_element_located((By.XPATH, "//img[contains(@src,\'https://{}/assets/plugins/images/planning.png\')]".format(url))))
    element = driver.find_element(By.XPATH, "//img[contains(@src,\'https://{}/assets/plugins/images/planning.png\')]".format(url))
    driver.execute_script("arguments[0].click();", element)
    
    driver.find_element(By.NAME, "user_name").send_keys(username)
    driver.find_element(By.NAME, "user_password").send_keys(password)
    driver.find_element(By.XPATH, "//button[@type=\'submit\']").click()

    WebDriverWait(driver, 3).until(expected_conditions.presence_of_element_located((By.XPATH, "//div[@id='page-wrapper-portal']/div[2]/div/div/div/div[2]/div[2]/div/i")))
    element = driver.find_element(By.XPATH, "//div[@id='page-wrapper-portal']/div[2]/div/div/div/div[2]/div[2]/div/i")
    driver.execute_script("arguments[0].click();", element)
    currentURL = driver.current_url
    if(currentURL.find("dashboard") > 0):
      status = 402
    else:
      status = 401
  except configparser.NoOptionError:
    set_file()
    write_file()
    status = 404
  return status
#sys.stdout = orig_stdout
#f.close()