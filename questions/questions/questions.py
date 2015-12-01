from pip._vendor import requests
import xlrd
import json
import configparser

# Parsing .ini file
config = configparser.ConfigParser()
config.read("C:/Users/mana/Documents/Visual Studio 2013/Projects/questions/questions/config.ini")
url = config.get("Environment", 'url')
path = config.get("Environment", 'path')
token = config.get("Environment", 'token')

# Read xls file
rb = xlrd.open_workbook(path,formatting_info=True)
sheet = rb.sheet_by_index(0)

# Constants
headers = {'Content-Type': 'application/json', 'x-wsse' : token}
passed = 0
failed = 0

for rownum in range(sheet.nrows):
    cell = sheet.cell_value(rownum, 0)
    r = requests.post(url, headers=headers, json={'text': cell })
    if r.status_code == 200:
        passed+=1
    else:
        failed+=1
        print ('Row number:', rownum)
        print(r.status_code, r.reason, r.text)

print(25*'.')
print("Created questions count:", passed)
print("Failed questions count:", failed)