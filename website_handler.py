# %% import modules
# selenium elements
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# other modules
import getpass
import time

# scripts
import globals as g

# store user name
strUserName = 'ivan.zustiak@zurich.com'

# request password
strPassword = getpass('Provide your password:\n')

# set up the webdriver
objDriver = webdriver.Edge()

# open the login
objDriver.get('https://my.xperience.app/zurich/login.jsp')

# find user id field
objUserName = objDriver.find_element('id', 'user')

# type in the user name
objUserName.send_keys(strUserName)

# find password field
objPassword = objDriver.find_element('id', 'pwd')

# type in user password
objPassword.send_keys(strPassword)

# find login button
objLoginButton = objDriver.find_element('id', 'loginButton')

# log in to the page
objLoginButton.click()

# open webpage for absence input
objDriver.get('https://my.xperience.app/zurich/web/timesystem?__mvcevent=absenceList&id=0')

# identify element with the absence types
objAbsenceType = objDriver.find_element('xpath', "//input[@class='ng-scope ng-isolate-scope custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left ui-autocomplete-input']")

# select home office absence
objAbsenceType.send_keys('home office')
time.sleep(1)
objAbsenceType.send_keys(Keys.TAB)
objAbsenceType.send_keys(Keys.ENTER)

# identify element with the start date
objDateStart = objDriver.find_element('xpath', "//input[@class='textbox_date info flapps-date aw-calendar hasDatepicker' and @name='startDate']")

objDateStart.send_keys('1/3/2023')

time.sleep(1)

# identify element with the end date
# objDateEnd = objDriver.find_element('xpath', "//input[@class='textbox_date info flapps-date aw-calendar hasDatepicker' and @name='endDate']")
# objDateEnd.send_keys('1/3/2023')

# for safety reasons click to note field
objNotes = objDriver.find_element('id', 'absenceDetailFM_note')
objNotes.click()

# find button to submit
objSubmit = objDriver.find_element('id', 'absenceDetailFM_submitRequest')
print(objSubmit.get_attribute('value'))