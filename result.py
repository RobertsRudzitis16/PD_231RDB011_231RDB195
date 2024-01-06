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
time.sleep(3)
find = driver.find_element(By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/button[2]')
find.click()


wb = Workbook()
ws = wb.active
max_row=ws.max_row
ws["A1"] = "Movie title"
ws["B1"] = "Year"
ws["C1"] = "Length"
ws["D1"] = "Rating"
ws["E1"] = "Popularity"
ws["F1"] = "Actors"
ws["G1"] = "Director"
ws["H1"] = "Description"

for name in movies[:3]:
    find = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "suggestion-search")))
    find.send_keys(name)
    find = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ipc-metadata-list-summary-item__t")))
    find.click()

    year_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/ul/li[1]')))
    length_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/ul/li[3]')))
    rating_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[2]/div/div[1]/a/span/div/div[2]/div[1]/span[1]')))
    popularity_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[2]/div/div[1]/a/span/div/div[2]/div[3]')))
    actor_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[2]/div/ul/li[3]/div/ul')))
    director_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[2]/div/ul/li[1]/div/ul/li/a')))
    description_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/p/span[2]')))
    
    year = year_element.text
    length = length_element.text
    rating = rating_element.text
    popularity = popularity_element.text
    actor = actor_element.text
    director = director_element.text
    description = description_element.text

    ws.append([name, year, length, rating, popularity, actor, director, description])




wb.save("movies.xlsx")    
input()