# Pytube-excel-list-downloader  
Input a playlist in excel and simply download all video to a folder from YouTube  
    
This script is a expansio for pytube library   
https://github.com/nficano/pytube

Dependency (all library include in this repository):
1. python3
2. pytube     (https://github.com/nficano/pytube)
3. xlrd       (https://github.com/python-excel/xlrd)
  
Support:  
1. Input a excel (.xls) as download list    
   - This feature is very practical, you can just copy urls from Youtube playlist and save urls to an .xls file  
   - You don't need to filter urls by yourself, this program can recognize youtube urls in the .xls file.  
   ![demo](https://raw.githubusercontent.com/SAM33/Pytube-excel-list-downloader/master/demo1.JPG)   
   ![demo](https://raw.githubusercontent.com/SAM33/Pytube-excel-list-downloader/master/demo2.JPG)   
2. Auto load/save download progress to continue your download  
3. Auto pause download session for a few seconds and restart again when youtube kill browse signature   
   
Usage:  
1. Browse youtube playlist by yourself, and copy it  
2. Select input playlist file (.xls format, not support .xlsx)   
3. Select output folder to save download video  
![demo](https://raw.githubusercontent.com/SAM33/Pytube-excel-list-downloader/master/demo3.JPG)   


