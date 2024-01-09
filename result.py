import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import Workbook, load_workbook 
import pandas as pd

print("Izvēlieties filmu saraksta kārtošanas metodi:")
print("1. Pēc reitinga")
print("2. Pēc gada")
print("3. Pēc popularitātes")
print("4. Pēc filmas garuma")
method = int(input("Ievadiet kārtošanas metodes numuru: "))
print()
print("Izvēlieties filmu saraksta kārtošanas virzienu:")
print("1. Augošā secībā")
print("2. Dilstošā secībā")
order = int(input("Ievadiet kārtošanas virziena numuru: "))

movies=[]

with open("movie_list.csv", "r", encoding="utf-8") as file:
    next(file)
    for line in file:
        movies.append(line)

service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)

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

data=[]

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

    ws.append([name, f"'{year}", length, rating, popularity, actor, director, description])



wb.save("movies.xlsx")  

df=pd.read_excel("movies.xlsx")
if method == 1:
    sort = "Rating"
elif method == 2:
    sort = "Year"
elif method == 3:
    sort = "Popularity"
elif method == 4:
    sort = "Length"

if order == 1:
     df_sorted=df.sort_values(by=sort, ascending=True)   
elif order == 2:
    df_sorted=df.sort_values(by=sort, ascending=False)

df_sorted.to_excel("movies_sorted.xlsx", index=False)

input()