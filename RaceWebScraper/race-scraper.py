from athlete_data import AthleteData
from athlete_data import fillExcelWorksheet
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook, load_workbook
from scraper_common import buildElitePointsArray
from crono4sport_live_scraper import ScrapeLiveCrono4SportFullRace
from crono4sport_scraper import ScrapeCrono4SportFullRace
from dorsalchip_scraper import ScrapeDorsalChipFullRace
import requests
import time
import os
import re


def scrapeFullRace(driver, url: str, elite_points, eventName: str, excelFilePath: str):
    print("We're going to scrape this url: " + url)
    if 'crono4sports.es/glive' in url:
        ScrapeLiveCrono4SportFullRace(driver, url, elite_points, eventName, excelFilePath)
    elif 'crono4sports.es' in url:
        ScrapeCrono4SportFullRace(driver, url, elite_points, eventName, excelFilePath)
    elif 'dorsalchip.es' in url:
        ScrapeDorsalChipFullRace(driver, url, elite_points, eventName, excelFilePath)
    else:
        print("Error!!! Unknown page to scrape: " + url)


#  configure the scraping 
print('Configuring scraping')
url = 'https://www.crono4sports.es/clasificacion/1664/'
raceFullName = 'Medieval Xtreme Race - Polop - 2024'
raceNum = 2
#url = 'https://www.crono4sports.es/glive/g-live.html?f=/carreras/1806-torrevieja.clax'
#raceFullName = 'Unbroken - Torrevieja - 2024'
#raceNum = 1
#url = 'https://www.dorsalchip.es/carrera/2024/2/18/SKULL_RACE.aspx'
#raceFullName = 'Skull Race - Torremolinos - 2024'
#raceNum = 3
#url = 'https://www.dorsalchip.es/carrera/2024/2/25/VI_The_Last_Race_.aspx#'
#raceFullName = 'The Last Race - Canillas de Aceituno - 2024'
#raceNum = 4
#url = 'https://www.crono4sports.es/clasificacion/1733/'
#raceFullName = 'Lion Race - Navas del Rey - 2024'
#raceNum = 5
#url = 'https://www.crono4sports.es/glive/g-live.html?f=/carreras/1699-llocnou.clax'
#raceFullName = 'Medieval Xtreme Race - Llocnou de San Jeromi - 2024'
#raceNum = 6
#url = 'https://www.crono4sports.es/glive/g-live.html?f=/carreras/1684-kongrace.clax'
#raceFullName = 'Kong Race - Polinya - 2024'
#raceNum = 7
isOCRSeries = False

# generate the file
currFolder = os.getcwd()
if isOCRSeries:
    directory = '\\data\\OCRSeries\\'
else:
    directory = '\\data\\LigaOCRA\\'
filename = str(raceNum) + " - " + raceFullName.replace(" ", "") + ".xlsx"
filePath = currFolder + directory + filename

# create the Selenium web driver
driver = webdriver.Chrome()

# build some info necessary for point calculation
elite_points = buildElitePointsArray()

# scrape the whole event
scrapeFullRace(driver, url, elite_points, raceFullName, filePath)

# close the browser
driver.quit()
