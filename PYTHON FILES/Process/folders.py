import utility
import os
from local import localPath

month = utility.getLastMonth()
year = utility.getLastMonthYear()
downloadedList = []
onlyZipList = []


def printDownloadedList():
    for file in downloadedList:
        print(f"{file[1]}   =>   {file[0]}")


def setDownloadedList():
    rfcFolders = os.listdir(f"{localPath}\\DESCARGAS")

    for rfcFolder in rfcFolders:
        path = f"{localPath}\\DESCARGAS\\{rfcFolder}\\{year}\\{month}"    
        files = os.listdir(path)

        for file in files:
            downloadedList.append([file, rfcFolder])
            downloadedList.append(file)

def getFielFiles(rfc):
    files = os.listdir(f"{localPath}\\FIEL\\" + rfc)
    dotcer = f"{localPath}\\FIEL\\" + rfc + "\\" + [f for f in files if f.endswith(".cer")][0]
    dotkey = f"{localPath}\\FIEL\\" + rfc + "\\" + [f for f in files if f.endswith(".key")][0]
    return dotcer, dotkey

def getImssFile(rfc):
    files = os.listdir(f"{localPath}\\IMSS\\" + rfc)
    dotpfx = f"{localPath}\\IMSS\\" + rfc + "\\" + [f for f in files if f.endswith(".pfx")][0]
    return dotpfx