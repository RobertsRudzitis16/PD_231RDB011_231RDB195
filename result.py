import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from openpyxl import Workbook, load_workbook 

service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)

movies=[]

with open("movie_list.csv", "r", encoding="utf-8") as file:
    next(file)
    for line in file:
        movies.append(line)

url = "https://www.imdb.com/"
driver.get(url)
time.sleep(2)

for name in movies:
    find=driver.find_element(By.ID, "suggestion-search")
    find.send_keys(name)

    find=driver.find_element(By.ID, "suggestion-search-button")
    find.click()
    time.sleep(2)

    find=driver.find_element(By.CLASS_NAME, "ipc-metadata-list-summary-item ipc-metadata-list-summary-item--click find-result-item find-title-result")
    find.click()
    
input()