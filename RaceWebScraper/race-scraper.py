from athlete_data import AthleteData
from athlete_data import fillExcelWorksheet
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook, load_workbook
from crono4sport_scraper import ScrapeCrono4SportFullRace
import requests
import time
import os
import re
from scraper_common import buildElitePointsArray


def scrapeFullRace(driver, url: str, elite_points, eventName: str, excelFilePath: str):
    print("We're going to scrape this url: " + url)
    if 'crono4sports.es' in url:
        ScrapeCrono4SportFullRace(driver, url, elite_points, eventName, excelFilePath)
    else:
        print("Error!!! Unknown page to scrape: " + url)


#  configure the scraping 
url = 'https://www.crono4sports.es/glive/g-live.html?f=/carreras/1806-torrevieja.clax'
currFolder = os.getcwd()
filePath = currFolder + '\\data\\Unbroken_Torrevieja_2024-02-17.xlsx'

# create the Selenium web driver
driver = webdriver.Chrome()

# build some info necessary for point calculation
elite_points = buildElitePointsArray()

# scrape the whole event
scrapeFullRace(driver, url, elite_points, 'Unbroken Torrevieja 2024', filePath)

# close the browser
driver.quit()
