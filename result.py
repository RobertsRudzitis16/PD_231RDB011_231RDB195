import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
find = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/button[2]')
find.click()

for name in movies:
    find = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "suggestion-search")))
    find.send_keys(name)
    find = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ipc-metadata-list-summary-item__t")))
    find.click()

wb = Workbook()
ws = wb.active
ws["A1"] = "Movie title"
ws["B1"] = "Year"
ws["C1"] = "Length"
ws["D1"] = "Rating"
ws["E1"] = "Popularity"
ws["F1"] = "Actors"
ws["G1"] = "Director"
ws["H1"] = "Description"
wb.save("movies.xlsx")
    
input()