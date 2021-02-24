__author__ = 'tatsuya'
import configparser
import glob
import json
import os
import pdfkit
import sys
import time
import xlrd
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

import LOGIN as login
import LOGOUT as logout

options = {
    'log-level':'error',
    'page-size':'A4',
    'dpi': 500,
    'orientation': 'landscape'
}
configur = configparser.ConfigParser()
#orig_stdout = sys.stdout
#f = open('info.API.txt', 'w+')
#sys.stdout = f
#Get semua file excel
try:
  folder = configur.get('api', 'folder')
except Exception:
  folder = "DATA_RKA"
file = glob.glob("{}/[!_][!~$]*.xlsx".format(folder))
#Fungsi Buka File Config
def write_file():
    configur.write(open('config.ini', 'w'))
#Fungsi Generate File Config
def set_file():
  configur['rka'] = {'folder': folder,'output': 'OUTPUT_RKA', 'indexFileBegin': '0', 'limitFilePerFolder': len(file)}
  configur['rka_files'] = {}
#Fungsi Generate Metadata dari File yang akan diproses
def set_file_index(index, mulai, batas, status):
  configur.set('rka_files', "filename-{}".format(index), '{}'.format(file[index]))
  configur.set('rka_files', "begin-{}".format(index), '{}'.format(mulai))
  configur.set('rka_files', "start-{}".format(index), '1')
  configur.set('rka_files', "limit-{}".format(index), '{}'.format(batas))
  configur.set('rka_files', "complete-{}".format(index), status)
def set_session(status, session, currenturl):
  configur.set('data', "status", '{}'.format(status))
  configur.set('data', "session", '{}'.format(session))
  configur.set('data', "currenturl", '{}'.format(currenturl))

#Cek keberadaan file config
if not os.path.exists('config.ini'):
  set_file()
  write_file()
else:
  configur.read('config.ini')

#Definisi konfigurasi file yang akan diproses
try:
  begin = configur.getint('rka', 'indexfilebegin')
  limit = configur.getint('rka', 'limitfileperfolder')
  status = configur.getint('data', 'status')
  fout = configur.get('rka', 'output')
except Exception:
  print("RKA Config Generated!")
  set_file()
  write_file()
  exit()

while status != 200:
  try:
    driver = webdriver.Chrome()
    baca = login.login(driver)
    print("Berhasil Login!\nKode: {}".format(baca))
    status = baca
    curl = driver.command_executor._url
    session_id = driver.session_id
    set_session(status, session_id, curl)
    write_file()
    exit()
  except Exception as err:
    #Jika terjadi kesalahan tampilkan diconsole dan lanjutkan
    driver.close()
    print("Gagal Login")
    status = 401
    set_session(status, "", "")
    write_file()
    continue

try:
  curl = configur.get('data', 'currenturl')
  session_id = configur.get('data', 'session')
  driver = webdriver.Remote(command_executor=curl,desired_capabilities={})
  driver.close()
  driver.session_id = session_id
  actionChains = ActionChains(driver)
except Exception as err:
  #Jika terjadi kesalahan tampilkan diconsole dan keluar
  print("Error")
  print(err)
  exit()

#Pembacaan file
for a in range(begin,limit):
  try:
    #Definisi nilai default dari metadata file yang akan diproses
    nama = "-"
    fbegin = 1
    flimit = 1
    process = "False"
    #Membuka file excel yang diproses
    workbook = xlrd.open_workbook(file[a])
    worksheet = workbook.sheet_by_index(0)
    #Check metadata dari Config
    nama = configur.get('rka_files', "filename-{}".format(a))
    fbegin = configur.getint('rka_files', "begin-{}".format(a))
    flimit = configur.getint('rka_files', "limit-{}".format(a))
    process = configur.get('rka_files', "complete-{}".format(a))
    #print("posisi: {} namafile: {} process {}".format(a,nama,process))
    #Check Status pemrosesan file
    if(process.lower() == "false"):
      #Blok untuk memproses data
      #print("Informasi excel fbegin: {} flimit: {} process TRUE".format(fbegin,flimit))
      counter = 0
      #Membaca data dari file excel
      path_wkthmltopdf = 'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
      config_pdf = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
      for b in range(fbegin,flimit):
        try:
          alamat = str(worksheet.cell(b, 31).value)
          nomor = str(worksheet.cell(b, 0).value)
          belanja = str(worksheet.cell(b, 40).value)
          id = str(worksheet.cell(b, 2).value)
          url = ("https://bandung.sipd.kemendagri.go.id/daerah/main/plan/belanja/2021/{}".format(alamat))
          driver.get(url)
          #Check ukuran file
          json_content = driver.page_source
          dirname = "D:\\Bandung\\SIPD\\SIPD_Automation\\{}\\SETDA".format(fout)
          if not os.path.exists(dirname):
            os.makedirs(dirname)
          dirhtml = "{}\\HTML".format(dirname)
          if not os.path.exists(dirhtml):
            os.makedirs(dirhtml)
          filename = '{}\\RKA.2021.{}.{}.pdf'.format(dirname, nomor, belanja)
          filehtml = '{}\\HTML\\RKA.2021.{}.{}'.format(dirhtml, nomor, belanja)
          pdfkit.from_string(json_content, filename, options=options, configuration=config_pdf)
          configur.set('rka_files', "begin-{}".format(a), '{}'.format(b))
          print("#{}. File Created: {}.{}.pdf".format(b, nomor, belanja))
          #driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
          #driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + 'w')
          #print("Informasi kolom-{} col0: {} col1: {} col2: {}".format(b,col0,col1,col2))
          #Tracking progress dan disimpan ke file config
          counter = counter + 1
          fbegin = b
          set_file_index(a,fbegin,flimit,process)
          write_file()
          with open("{}.html".format(filehtml), 'w+') as f:
            f.write(json_content)
        except Exception as err:
          #Jika terjadi kesalahan tampilkan diconsole dan lanjutkan
          print(err)
          pass
      #Cek status dari data yang diproses
      if(counter>0 or fbegin >= flimit):
        process = "True"
        set_file_index(a,fbegin,flimit,process)
        write_file()
        baca = logout.logout(driver)
        set_session(baca, "", "")
        print("Berhasil Logout!\nKode : {}".format(baca))
    elif(process.lower() == "true"):
      #Blok ketika data sudah diproses sebelumnya
      counter = 0
      #print("Informasi excel fbegin: {} flimit: {} file sudah diproses (COMPLETE)".format(fbegin,flimit))
    else:
      #Blok ketika data tidak valid saat dijalankan sebelumnya
      counter = 0
      #print("Informasi excel fbegin: {} flimit: {} process ERROR {}".format(fbegin,flimit, process))
  except configparser.NoOptionError:
    #Blok ketika metadata belum tersedia di file config
    flimit = worksheet.nrows
    set_file_index(a,fbegin,flimit,process)
    print("Metadata Generated!")
    #print("Filename: {}\nBegin : {}\nLimit : {}\nProcess : {}".format(file[a],fbegin,flimit, process))
    continue
  except Exception as err:
    #Blok ketika terjadi kesalahan
    print(err)
    process = "Error"
    set_file_index(a,fbegin,flimit,process)
    break

write_file()
#sys.stdout = orig_stdout
#f.close()