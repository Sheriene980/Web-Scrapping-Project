import os
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from urllib.request import urlopen
from shutil import copyfileobj

import pytube

# Change files here by keying in the file name
file_name = "Question_1_dataset_Automation_Specialist"

# Changing file to csv
file_imported = pd.read_excel(f"{file_name}.xlsx")
csv_file = file_imported.to_csv (f"{file_name}", index = None, header=True)
data = pd.read_csv(f"{file_name}.csv")

# Downloading images
def download_images(data_item):
    all_images_link = driver.find_elements(By.XPATH,"//meta[@property='og:image']")
    for image_link in all_images_link:
        images = image_link.get_attribute("content")
        with urlopen(images) as in_stream, open(data_item, 'wb') as out_file:
            copyfileobj(in_stream, out_file)
    print(f'Downloaded {len(all_images_link)} images')

 # Downloading videos
def download_videos(SKU_file):
    try:
        el = driver.find_element(By.XPATH, '//*/div[@id="playButton"]')
        webdriver.ActionChains(driver).move_to_element(el).click(el).perform()
        print('Video detected')

        WebDriverWait(driver, 10)
        vid_player = driver.find_element(By.XPATH,'//*/iframe[@id="videoPlayer"]')
        yt_link = vid_player.get_attribute('src')
        driver.get(yt_link)

        vid_link = driver.find_element(By.XPATH,'//head/link[@rel="canonical"]')
        vid_src = vid_link.get_attribute('href')
        print(vid_src)

        yt = pytube.YouTube(vid_src)
        video = yt.streams.first()
        video.download(SKU_file)
        print('Successfully downloaded video!')

    except:
        print(f'No video detected')   

# File to store all SKU files
try:
    os.mkdir('product_media/')
except OSError as error:
    print(error)

s=Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)

not_found = []

for data_item in data["Sku"]:

    # Create file paths based on SKUs to download product media
    SKU_file = f'product_media/{data_item}/'
    try:
        os.mkdir(SKU_file)
    except OSError: # Error message when SKU has been searched before
        print(f'Warning: File for SKU {data_item} has been made')

    zalora_site = f"https://www.zalora.com.my/catalog/?q={data_item}"
    print(f'Searching for SKU {data_item}: {zalora_site}')

    with webdriver.Chrome(service=s) as driver:
        driver.get(zalora_site)

        # Search for products by SKU
        SKU_searches = driver.find_elements(By.XPATH,'//div/div/ul/li/a[@href]')
        for items in SKU_searches:
            link = items.get_attribute('href')
            # print(link)
            if link != None and link[-5:] == '.html':
                # print(link)
                break

        if link[-5:] != '.html':
            print(f'Note: SKU {data_item} is not found.')
            not_found.append(data_item)
            driver.close()
            continue

        driver.get(link)

        # Downloading images
        download_images(data_item)

         # Downloading videos
        download_videos(SKU_file)

        driver.close()

print(f'========= Product Media Download Report =========')
print('Out of',len(data),'SKUs, \n',len(not_found),'SKUs not found: ',*not_found)   