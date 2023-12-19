# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 13:51:27 2023

@author: monster
"""

#import datetime  
from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#from openpyxl import load_workbook
#from openpyxl import Workbook
import pandas as pd


driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://www.google.com")
aramakutusu = driver.find_element(By.NAME,"q")
aramakutusu.send_keys("döviz")
aramakutusu.send_keys(Keys.ENTER)
sleep(2)
sayfayagiris = driver.find_element(By.CSS_SELECTOR,"#rso > div.hlcw0c > div > div > div > div > div > div > div > div.yuRUbf > div > span > a > h3")
sayfayagiris.click()
driver.get("https://kur.doviz.com")
driver.execute_script("window.scrollBy(0,750)", "")
sleep(5)
kaynak = driver.page_source
soup = BeautifulSoup(kaynak, "html.parser")


"""
     
try:
    workbook = load_workbook("veriler.xlsx")
except FileNotFoundError:
    workbook = Workbook()
    
   
sayfa_adi = "Sheet" 
if sayfa_adi in workbook.sheetnames:
    workbook.remove(workbook[sayfa_adi])  
  
now = datetime.datetime.now() 
sayfa_adi = f"Tarih {now.year}{now.month:02d}{now.day:02d}-{now.hour:02d}{now.minute:02d}" # Mevcut tarih bazında sayfa adı oluşturur. Örneğin: Veriler20230102
if sayfa_adi in workbook.sheetnames: 
    workbook.remove(workbook[sayfa_adi]) 

sayfa = workbook.create_sheet(title=sayfa_adi)
sayfa["A1"] = "Döviz Kurları"
sayfa["B1"] = "Alış Fiyatı"
sayfa["C1"] = "Satış Fiyatı"



"""


metinler = soup.find_all("div" , attrs={"class":"market-data"})

liste = []

for metin in metinler:
    baslık = metin.find("div" , attrs={"class": "item"})
    liste.append(baslık.text)
    
    baslık = metin.find("a" , attrs={"href": "https://kur.doviz.com/serbest-piyasa/euro"})
    liste.append(baslık.text)
    
    baslık = metin.find("a" , attrs={"href": "https://kur.doviz.com/serbest-piyasa/amerikan-dolari"})
    liste.append(baslık.text)
    
    baslık = metin.find("a" , attrs={"href": "https://kur.doviz.com/serbest-piyasa/sterlin"})
    liste.append(baslık.text)
    
    baslık = metin.find("a" , attrs={"href": "https://borsa.doviz.com/endeksler/xu100-bist-100"})
    liste.append(baslık.text)
    
    
    
    
  

#workbook.save("veriler.xlsx")  
veritabanı = pd.DataFrame(liste)  
veritabanı.to_excel("veriler.xlsx")


#for veri in liste:
        
#    sayfa.append(veri)
   
    
     
  

print(driver.title)
print(driver.current_window_handle)
driver.switch_to.new_window("tab")
driver.get("https://kur.doviz.com/kapalicarsi")
print(driver.title)
print(driver.current_window_handle)
driver.execute_script("window.scrollBy(0,650)", "")
sleep(5)
print(driver.title)
print(driver.current_window_handle)
driver.switch_to.new_window("tab")
driver.get("https://altin.doviz.com")
driver.execute_script("window.scrollBy(0,565)", "")
sleep(5)
print(driver.title)
print(driver.current_window_handle)
driver.switch_to.new_window("tab")
driver.get("https://borsa.doviz.com")
driver.execute_script("window.scrollBy(0,770)", "")
sleep(5)
print(driver.title)
print(driver.current_window_handle)
driver.switch_to.new_window("tab")
driver.get("https://www.doviz.com/akaryakit-fiyatlari")
driver.execute_script("window.scrollBy(0,300)", "")
sleep(5)