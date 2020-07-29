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

options = {
    'log-level':'error'
}
configur = configparser.ConfigParser()
#Fungsi Buka File Config
def write_file():
  configur.write(open('config.ini', 'w'))
#Fungsi Generate File Config
def set_session(status, session, currenturl):
  configur.set('data', "status", "{}".format(status))
  configur.set('data', "session", "{}".format(session))
  configur.set('data', "currenturl", "{}".format(currenturl))

#Cek keberadaan file config
if not os.path.exists('config.ini'):
  set_session("501", "", "")
  write_file()
else:
  configur.read('config.ini')
  set_session("501", "", "")
  write_file()

#Function Logout:
def logout(driver):
  #Definisi konfigurasi file yang akan diproses
  try:
    url = configur.get('data', 'url')
    driver.get("https://{}/daerah/main/plan/logout/0".format(url))
    currentURL = driver.current_url
    if(currentURL.find("landing") > 0):
      status = 501
    else:
      status = 402
  except configparser.NoOptionError:
    status = 404
  set_session("{}".format(status), "", "")
  write_file()
  driver.close()
  return status