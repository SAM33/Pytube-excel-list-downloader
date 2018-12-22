import sys
import os
import xlrd
import json
from pytube import YouTube
import time
from random import randint
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
download_list_file = askopenfilename() # show an "Open" dialog box and return the path to the selected file
save_to_folder = askdirectory()
if download_list_file and save_to_folder:
    print("Download list file: %s" % download_list_file)
    print("Save to folder: %s" % save_to_folder)
else:
    print("download_list_file and save_to_folder not correct")
    raise
xls_path = download_list_file
progress_save = "%s.sav.json" % xls_path
book = xlrd.open_workbook(download_list_file)
mainData_sheet = book.sheet_by_index(0)
watch_url = []
for row in range(1, mainData_sheet.nrows):
    rowValues = mainData_sheet.row_values(row, start_colx=0, end_colx=8)
    company_name = rowValues[0]
    link = mainData_sheet.hyperlink_map.get((row, 0))
    url = '(No URL)' if link is None else link.url_or_path
    if 'watch?v=' in url:
        watch_url.append(url.split('&list=')[0])
download_urls = []
download_ok = []
jobj = {}
try:
    print("load save progress...")
    f = open(progress_save)
    jdata = f.read()
    f.close()
    jobj = json.loads(jdata)
    download_ok = jobj
    print("continue download... %d of %d" % (len(download_ok),len(watch_url)))
except:
    print("save progress not exist")
    print("or save progress file broken")
for i in range(0,len(watch_url)):
    download_urls.append(watch_url[i])
download_path = save_to_folder
idx = 0
for link in download_urls:
    try:
        if link in download_ok:
            idx += 1
            continue
        print("Download %d of %d" % (idx,len(watch_url)))
        yt = YouTube(link)
        dl_stream = yt.streams.filter(
            progressive=True, subtype='mp4',
        ).order_by('resolution').desc().first()
        print(link)
        dl_stream.download(download_path)

        download_ok.append(link)
        print("save progress %d of %d" % (idx, len(watch_url)))
        jdata = json.dumps(download_ok)
        f = open(progress_save,"w")
        f.write(jdata)
        f.close()
        idx += 1
    except Exception as e:
        print("This session has been block...")
        sleep_time = 50+randint(0, 50)
        print("Stop dowonload for %d seconds" % sleep_time)
        time.sleep(sleep_time)
        print("Try again...")
