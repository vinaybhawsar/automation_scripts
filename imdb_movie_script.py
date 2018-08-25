# -*- coding: utf-8 -*-
import os
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys  
from selenium.webdriver.chrome.options import Options  

chromedriver = "/home/user/Downloads/chromedriver"
os.environ["webdriver.chrome.driver"] = chromedriver

chrome_options = Options()
chrome_options.add_argument("--headless") # Runs Chrome in headless mode.
chrome_options.add_argument('--no-sandbox') # # Bypass OS security model
chrome_options.add_argument('start-maximized')
chrome_options.add_argument('disable-infobars')
chrome_options.add_argument("--disable-extensions")
driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chrome_options)

driver = webdriver.Chrome(chromedriver)

def write_to_file(driver, workbook):
	"""Function to write data in excel sheet"""
	genre_list = driver.find_elements_by_class_name('subnav_item_main')
	genre_links = {}
	for genre in genre_list:
		genre_links[genre.text] = genre.find_element_by_tag_name("a").get_attribute("href")

	try:
		for genre in genre_links.keys():
			print genre
			worksheet = workbook.add_worksheet(genre)
			genre_link = genre_links[genre]
			driver.get(genre_link)
			row = 0
			while True:
				t2 = driver.find_element_by_class_name('lister-list') 
				t2 = t2.find_elements_by_class_name('lister-item')
				col = 0
				worksheet.write(row, col, "Rating")
				worksheet.write(row, col + 1, "Name")
				worksheet.write(row, col + 2, "Genre")
				for t in t2:
					col = 0
					try:
						t21 = t.find_element_by_class_name('lister-item-header')
						t22 = t.find_element_by_class_name('ratings-bar')
						t23 = t.find_element_by_class_name('genre')

						worksheet.write(row, col, t22.text.split("Rate this")[0].strip())
						col += 1
						worksheet.write(row, col, "".join(t21.text.split(".")[1:]).strip())
						col += 1
						worksheet.write(row, col, t23.text.strip())
						row += 1
						print t21.text, t22.text.split("Rate this")[0], t23.text
					except Exception as e:
						print genre , " Issue"
						pass
				try:
					t3 = driver.find_element_by_link_text('Next »')
					t3.click()
				except Exception as e:
					print genre , " Complete"
					break
	except Exception as e:
		print e
		pass

	workbook.close()

workbook = xlsxwriter.Workbook('Movies List.xlsx')
driver.get("https://www.imdb.com/chart/top")
write_to_file(driver, workbook)

workbook = xlsxwriter.Workbook('TV Show List.xlsx')
driver.get("https://www.imdb.com/chart/toptv")
write_to_file(driver, workbook)

driver.quit()