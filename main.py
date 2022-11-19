#   For exe version validation
import win32api
#   For comparing official webpage verison
from selenium import webdriver
from selenium.webdriver.common.by import By
import wget
import subprocess, os, time
import psutil

#   Get file version of the exe file
#   https://stackoverflow.com/questions/68774795/grabbing-full-file-version-of-an-exe-in-python
def get_file_version(path):
    info = win32api.GetFileVersionInfo(path, '\\')
    ms = info['FileVersionMS']
    ls = info['FileVersionLS']
    return (str(win32api.HIWORD(ms)) + "." + str(win32api.LOWORD(ms))
     + "." + str(win32api.HIWORD(ls))) 

def checkIfProcessRunning(processName):
    #Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

print("Checking for latest version...")
#   Download URL
downloadPage = "https://picpick.app/en/download/free"
#   Use Edge
driver = webdriver.Edge()
driver.minimize_window()
driver.get(downloadPage)
#   Find the link
exeURL = driver.find_element("id", 'dwstart')
exeURL = exeURL.get_attribute('href')
driver.close()

latestVersion = exeURL.split("/")[3]

currentVersion = get_file_version("C:\Program Files (x86)\PicPick\picpick.exe") 

if(currentVersion != latestVersion):
    response = wget.download(exeURL, exeURL.split("/")[4])
    if(os.path.exists(exeURL.split("/")[4])):
        os.system(exeURL.split("/")[4] + " /S")
        #   If the installer is not running we are good to delete the installer
        while(True):
            if not (checkIfProcessRunning(exeURL.split("/")[4].split(".")[0])): break
        os.remove(exeURL.split("/")[4])
        

    
    


