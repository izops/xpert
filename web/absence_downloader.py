# %% import modules
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import datetime
import sys

sys.path.append('c:/repositories/emea_oth_xpert')
import general.global_constants as g
import general.general_functions as ggf
import web.common_web_fucntions as wcf

# %% define functions to webscrape data
def lstDownloadData(
    pstrUserName,
    pstrPassword,
    pstrDateFrom,
    pstrDateTo,
    pstrAbsenceType
):
    """Download absence data from xperience in the given period.
    
    Inputs:
        - pstrUserName - username for xperience login
        - pstrPassword - password for xperience login
        - pstrDateFrom - start date of absence filtering, YYYYMMDD string
        - pstrDateTo - end date of absence filtering, YYYYMMDD string
        - pstrAbsenceType - absence type to scrape

    Outputs:
        - lstAbsences - list of all data scraped from absence list
    """
    # initialize webdriver
    objDriver = webdriver.Edge()

    # login to the xperience
    blnLogin = wcf.blnLogin(objDriver, pstrUserName, pstrPassword)

    if blnLogin:
        # load page with the absence list
        objDriver.get(g.STR_URL_LIST_ABSENCES)

        # initialize list of absences
        lstAbsences = []

        # set the filter start date
        objDateStart = objDriver.find_element('id', g.STR_ELEMENT_ID_DATE_FROM)
        objDateStart.send_keys(pstrDateFrom)

        # set the filter end date
        objDateEnd = objDriver.find_element('id', g.STR_ELEMENT_ID_DATE_TO)
        objDateEnd.send_keys(pstrDateTo)

        # open status dropdown
        objStatus = objDriver.find_element('xpath', g.STR_ELEMENT_XPATH_STATUS)
        objStatus.click()

        # wait for the javascript to kick in
        time.sleep(1)

        # set the relevant status
        objApproved = objDriver.find_element(
            'id',
            g.STR_ELEMENT_ID_CHECKBOX_APPROVED
        )
        objApproved.click()

        objClosed = objDriver.find_element(
            'id',
            g.STR_ELEMENT_ID_CHECKBOX_CLOSED
        )
        objClosed.click()

        # close the status dropdown
        objStatus.click()

        # set absence type
        objType = objDriver.find_element('xpath', g.STR_ELEMENT_XPATH_TYPE)
        objType.send_keys(pstrAbsenceType)

        # wait for the javascript
        time.sleep(1)

        # confirm selection
        objType.send_keys(Keys.TAB)

        # confirm filtering
        objButton = objDriver.find_element(
            'id',
            g.STR_ELEMENT_ID_BUTTON_FILTER
        )
        objButton.click()

        # locate the table
        objTable = objDriver.find_element('xpath', g.STR_ELEMENT_XPATH_TABLE)

        # locate the table rows
        objRows = objTable.find_elements('tag name', 'tr')

        # loop through the table and save the values to a list
        for objRow in objRows:
            # initialize list of absence details
            lstDetails = []

            # locate all row cells
            objCells = objRow.find_elements('tag name', 'td')

            # store value from every cell of the row to a list
            for objCell in objCells:
                lstDetails.append(objCell.text)

            # save the details to the list of all absences
            lstAbsences.append(lstDetails)

    return lstAbsences