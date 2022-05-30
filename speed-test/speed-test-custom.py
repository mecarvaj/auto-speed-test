from openpyxl import Workbook
import openpyxl

import speedtest
from datetime import datetime


wifi  = speedtest.Speedtest()
try:
    workbook = openpyxl.load_workbook("speed-test.xlsx")

except:
    workbook = Workbook()

sheet = workbook.active

timestamp = datetime.now()
download = wifi.download()/1024.0**2
upload = wifi.upload()/1024.0**2


row = [timestamp, "{:.2f}".format(download), "{:.2f}".format(upload)]
sheet.append(row)
workbook.save(filename="speed-test.xlsx")
