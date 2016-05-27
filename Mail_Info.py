from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup as BS

import xlsxwriter


Browser = webdriver.Firefox()
Browser.get("https://mail.zyxel.com.tw/owa/auth/logon.aspx?replaceCurrent=1&url=https%3a%2f%2fmail.zyxel.com.tw%2fowa%2f")
Browser.find_element_by_id("username").send_keys(Username)
Browser.find_element_by_id("password").clear()
Browser.find_element_by_id("password").send_keys(Passowrd)
Browser.find_element_by_css_selector("input.btn").click()

workbook = xlsxwriter.Workbook('ZyMail.xlsx')
worksheet = workbook.add_worksheet()

#Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold' : True})

#Adjust the column width.
worksheet.set_column("A:A", 100)
worksheet.set_column("B:B", 23)
worksheet.set_column("C:C", 17.14)

#Start from the first cell below the headers.
row = 1
col = 0  

#Write some data header.


worksheet.write(0, 0, 'Title',bold)  
worksheet.write(0, 1, 'Sender',bold)  

soup=BS(Browser.page_source,"html.parser")

for ele in soup.select('.c3'):
	i = row % 2
	j = int(row /2)+1
	if i==1: 
		worksheet.write(j, 0, ele.text)
	else:
		worksheet.write(j-1, 2, ele.text)
	row=row+1
	
	
col=col+1
row = 1

for ele in soup.select('.c2'):
	i = row / 2	
	worksheet.write(i, col, ele.text)	
	row=row+1
	
Browser.close()

workbook.close()
