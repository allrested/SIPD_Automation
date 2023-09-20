# SIPD_Automation
 Web Automation using selenium written in Python

## Requirements

### Basic requirements:

 1. Python

	```
	$ sudo apt-get install python-pip
	```

 2. Chromedriver

	```
    $ sudo apt-get update
	$ sudo apt-get install -y unzip xvfb libxi6 libgconf-2-4
    $ wget https://chromedriver.storage.googleapis.com/2.41/chromedriver_linux64.zip
    $ unzip chromedriver_linux64.zip
    $ 
    $ sudo mv chromedriver /usr/bin/chromedriver
    $ sudo chown root:root /usr/bin/chromedriver
    $ sudo chmod +x /usr/bin/chromedriver
	```

## Setup

```
$ pip install selenium
```

## Test run

```
$ git clone https://github.com/allrested/SIPD_Automation.git
$ cd SIPD_Automation
```

## Run Data Crawler run

```
$ python API.py
$ API Config Generated!
$
$ python API.py
$ DevTools listening on ws://127.0.0.1:1234/
$ Berhasil Login!
$ Kode: 200
$
$ python API.py
$ DevTools listening on ws://127.0.0.1:1234/
$ Metadata Generated!
$
$ python API.py
$ DevTools listening on ws://127.0.0.1:1234/
$ #1. File Created: VISI.xlsx
$ #2. File Created: MISI.xlsx
$ #3. File Created: TUJUAN.xlsx
$ Berhasil Logout!
$ Kode : 501
$
```
