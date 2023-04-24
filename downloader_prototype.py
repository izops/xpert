# %% import modules
# selenium elements
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import website_handler as w
import globals as g


# %% constants
STR_URL_ABSENCES = 'https://my.xperience.app/zurich/web/timesystem?__mvcevent=absenceList'

STR_ELEMENT_ID_FROM = 'absenceRequestListTM_fd_dateFrom'
STR_ELEMENT_ID_TO = 'absenceRequestListTM_fd_dateTo'
STR_ELEMENT_XPATH_STATUS = '//input[@class = "ui-widget ui-widget-content ui-corner-left ui-autocomplete-input"]'
STR_ELEMENT_XPATH_TYPE = '//input[@class = "ng-scope ng-isolate-scope custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left ui-autocomplete-input"]'

STR_ELEMENT_CHECKBOX_ID_APPROVED = 'checkfia_status1'
STR_ELEMENT_CHECKBOX_ID_CLOSED = 'checkfia_status3'

STR_ELEMENT_ID_BUTTON = 'absenceRequestListTM__find'

STR_ELEMENT_XPATH_TABLE = '//tbody[@class = "data"]'

# %%

# set username and password
strUsername = w.strGetUserName()
strPwd = '[po[po{PO0'

# initialize webdriver
objDriver = webdriver.Edge()

# login
blnLogin = w.blnLogin(objDriver, strUsername, strPwd)

# load absence list
objDriver.get(STR_URL_ABSENCES)

# find from
objFrom = objDriver.find_element('id', STR_ELEMENT_ID_FROM)
objFrom.send_keys('1.1.2020')


# find to
objTo = objDriver.find_element('id', STR_ELEMENT_ID_TO)
objTo.send_keys('1.1.2023')

# set status
objStatus = objDriver.find_element('xpath', STR_ELEMENT_XPATH_STATUS)
objStatus.click()

time.sleep(1)

objApproved = objDriver.find_element('id', STR_ELEMENT_CHECKBOX_ID_APPROVED)
objApproved.click()

objClosed = objDriver.find_element('id', STR_ELEMENT_CHECKBOX_ID_CLOSED)
objClosed.click()

# close the menu
objStatus.click()

# set type
objType = objDriver.find_element('xpath', STR_ELEMENT_XPATH_TYPE)
objType.send_keys('Vacation')

time.sleep(1)
objType.send_keys(Keys.TAB)

# confirm
objButton = objDriver.find_element('id', STR_ELEMENT_ID_BUTTON)
objButton.click()

# read the table
objTable = objDriver.find_element('xpath', STR_ELEMENT_XPATH_TABLE)
objRows = objTable.find_elements('tag name', 'tr')

# loop through the table and print the values from it to console
for objRow in objRows:
    objCells = objRow.find_elements('tag name', 'td')

    for objCell in objCells:
        print(objCell.text)