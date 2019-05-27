# smartstore-automate
automation for naver smart store


# requirements

* selenium
* openpyxl
* xlwt
* requests

# Installation

```
sudo sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'
wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | sudo apt-key add -
sudo apt-get update
sudo apt-get install google-chrome-stable

sudo apt-get install python3-venv
python3 -m venv venv
source venv/bin/activate
pip install wheel
pip install -r requirements.txt
```

# Files
## secret.py

```
executable_path = 'drivers/linux/chromedriver-74'

chrome_profile_path = 'chrome profile/settings path(absolute)'

chrome_headless = True

download_path = 'download path(absolute) for temporary excel files'

username = 'Naver ID'

password = 'Naver Password'

batch_excel = 'batch.xls'

# production
api_url = 'REST API URL'

api_token = 'REST API TOKEN'
```
## autosend.sh

```
#!/bin/bash

cd /home/xxx
source venv/bin/activate
cd smartstore-automate
python smartstore.py
```

## crontab run script every 5 minutes

```
*/5 * * * * /bin/bash /home/xxx/autosend.sh >/home/xxx/autosend.log 2>&1
```

# TODO

* send request (async)
