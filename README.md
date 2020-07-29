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
$ python LOGIN.py
$ python API.py
```