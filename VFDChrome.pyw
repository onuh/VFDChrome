import json, os, sys, psutil
from configparser import ConfigParser
import subprocess
from subprocess import check_output
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
import chromedriver_autoinstaller_fix
import time
from random import randint
from shutil import copyfile
from datetime import datetime
import ctypes


def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def chech_chrome():
    running_name_of_process = []
    for process in psutil.process_iter():
        if process.name() == "chrome.exe":
            running_name_of_process.append(process.name())

    running = running_name_of_process.count("chrome.exe")
    if running >= 1:
        time.sleep(0.5)
        close_chrome_process = check_output('wmic process where name="chrome.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))


def verification():
    running_name_of_process = []
    for process in psutil.process_iter():
        if process.name() == "VFDChrome.exe":
            running_name_of_process.append(process.name())

    running = running_name_of_process.count("VFDChrome.exe")
    if running > 1:
        time.sleep(0.5)
        close_chrome_process1 = check_output('wmic process where name="chrome.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        close_chrome_process2 = check_output('wmic process where name="VFDChrome.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        close_chrome_process3 = check_output('wmic process where name="chromedriver.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        sys.exit()
    else:
        close_chrome_process1 = check_output('wmic process where name="chrome.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        close_chrome_process3 = check_output('wmic process where name="chromedriver.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        return


def home_page_config():
    if os.path.exists("c:\\VFDBoxChrome"):
        config = ConfigParser()
        try:
            config.read("c:\\VFDBoxChrome\\config.ini")
            configured_hompage = config.get("main", "homepage")
            return configured_hompage
        except Exception as no_hompage:
            config.add_section("main")
            config.set("main", "homepage", "https://google.com")
            with open("c:\\VFDBoxChrome\\config.ini", "w") as f:
                config.write(f)
                return "https://google.com"

    if not os.path.exists("c:\\VFDBoxChrome"):
        config = ConfigParser()
        config.add_section("main")
        config.set("main", "homepage", "https://google.com")
        with open("c:\\VFDBoxChrome\\config.ini", "w") as f:
            config.write(f)
            return "https://google.com"


def close_all_chrome_process():
    try:
        close_chrome_process = check_output('wmic process where name="chrome.exe" call terminate', shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
    except Exception as e:
        return


def write_log_file(log_text):
    data_log_path = "c:\\VFDBoxChrome\\logs"
    if not os.path.exists(data_log_path):
        os.makedirs(data_log_path)
    date = datetime.today().strftime("%Y-%m-%d")
    current_time = datetime.now().time()
    current_time = current_time.strftime("%I:%M %p")
    log_data = str(date) + " " + str(current_time) + " : " + log_text
    with open(data_log_path + "\\ProcessedJobs.log", "a") as mylog:
        mylog.write("\r")
        mylog.write(log_data)
        return


def monitor_download_folder(driver):
    path = "c:\\VFDBoxChrome\\Invoices"
    vfdpath = "c:\\ProgramData\\VFDBox\\InvoicingService\\InvoicesIn"
    if not os.path.exists(path):
        os.makedirs(path)
    if not os.path.exists(vfdpath):
        os.makedirs(vfdpath)
    # while True:
    try:
        driver_log = driver.get_log('driver')
        print(driver_log)
        if driver_log and driver_log[0]['message'] == 'Unable to evaluate script: no such window: target window already closed\nfrom unknown error: web view not found\n':
            print('exiting...')
            check_output("taskkill /f /im chromedriver.exe", shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
            sys.exit()
        time.sleep(1)
        file_name = randint(0, 999999999)
        for file in os.listdir(path):
            if file.endswith(".pdf"):
                try:
                    os.rename(path + "\\" + str(file), path + "\\receipt_" + str(file_name) + ".pdf")
                    copyfile(path + "\\receipt_" + str(file_name) + ".pdf", vfdpath + "\\receipt_" + str(file_name) + ".pdf")
                    os.remove(path + "\\receipt_" + str(file_name) + ".pdf")
                    write_log_file("receipt_" + str(file_name) + ".pdf sent to VFD for processing")
                except Exception as err:
                    write_log_file(str(err))
                    try:
                        os.remove(path + "\\receipt_" + str(file_name) + ".pdf")
                    except Exception as del_err:
                        write_log_file(str(del_err))
                        pass
    except Exception as err:
        print("Error Occured: ", str(err))
        ret1 = check_output("taskkill /f /im chromedriver.exe", shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
        sys.exit()
    

chrome_options = webdriver.ChromeOptions()
download_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), "Invoices")
appState = {'recentDestinations':[
  {'id':"Save as PDF",
   'origin':"local",
   'account':""}],
 'isHeaderFooterEnabled':False,
 'selectedDestinationId':"Save as PDF",
 'version':2}
profile = {'plugins.plugins_list':[
  {'enabled':False,
   'name':"Chrome PDF Viewer"}],
 'download.default_directory':download_dir,
 'profile.default_content_settings.popups':0,
 'download.extensions_to_open':"applications/pdf",
 'printing.print_preview_sticky_settings.appState':json.dumps(appState),
 'savefile.default_directory':"c:\\VFDBoxChrome\\Invoices"}
chrome_options.add_experimental_option("prefs", profile)
chrome_options.add_argument("--kiosk-printing")
chrome_options.add_argument("--kiosk")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

if __name__ == "__main__":
    print("Start run install driver")
    try:
        chromedriver_autoinstaller_fix.install()

    except Exception as e:
        Mbox('Connection Error', 'Could not connect to chromedriver update server', 1)
        sys.exit()
    close_all_chrome_process()
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.get(home_page_config())
    print(driver.get_log('driver'))
    print("loop started")
    while True:
        # chech_chrome()
        path = "c:\\VFDBoxChrome\\Invoices"
        vfdpath = "c:\\ProgramData\\VFDBox\\InvoicingService\\InvoicesIn"
        if not os.path.exists(path):
            os.makedirs(path)
        if not os.path.exists(vfdpath):
            os.makedirs(vfdpath)
        # while True:
        try:
            driver_log = driver.get_log('driver')
            print(driver_log)
            if driver_log and driver_log[0]['message'] == 'Unable to evaluate script: no such window: target window already closed\nfrom unknown error: web view not found\n':
                print('exiting...')
                check_output("taskkill /f /im chromedriver.exe", shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
                sys.exit()
            time.sleep(1)
            file_name = randint(0, 999999999)
            for file in os.listdir(path):
                if file.endswith(".pdf"):
                    try:
                        os.rename(path + "\\" + str(file), path + "\\receipt_" + str(file_name) + ".pdf")
                        copyfile(path + "\\receipt_" + str(file_name) + ".pdf", vfdpath + "\\receipt_" + str(file_name) + ".pdf")
                        os.remove(path + "\\receipt_" + str(file_name) + ".pdf")
                        write_log_file("receipt_" + str(file_name) + ".pdf sent to VFD for processing")
                    except Exception as err:
                        write_log_file(str(err))
                        try:
                            os.remove(path + "\\receipt_" + str(file_name) + ".pdf")
                        except Exception as del_err:
                            write_log_file(str(del_err))
                            pass
        except Exception as err:
            print("Error Occured: ", str(err))
            ret1 = check_output("taskkill /f /im chromedriver.exe", shell=True, stdin=(subprocess.PIPE), stderr=(subprocess.PIPE))
            sys.exit()
