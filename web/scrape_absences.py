# %% import modules
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import general.general_functions as ggf
import web.common_web_functions as wcf

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
        
        # split the date range to smaller ranges
        lstDateRanges = ggf.lstSplitDates(pstrDateFrom, pstrDateTo)
        
        # scrape data for each time subperiod
        for tplCurrentDate in lstDateRanges:
            # locate the filter start date
            objDateStart = objDriver.find_element(
                'id',
                g.STR_ELEMENT_ID_DATE_FROM
            )

            # clear the field contents
            objDateStart.clear()

            # convert the date to xperience format and type to field
            objDateStart.send_keys(
                ggf.strChangeDateFormat(tplCurrentDate[0], '.')
            )

            # locate the filter end date
            objDateEnd = objDriver.find_element(
                'id',
                g.STR_ELEMENT_ID_DATE_TO
            )

            # clear the field contents
            objDateEnd.clear()

            # convert the date to xperience format and type to field
            objDateEnd.send_keys(
                ggf.strChangeDateFormat(tplCurrentDate[1], '.')
            )

            # open status dropdown
            objStatus = objDriver.find_element(
                'xpath',
                g.STR_ELEMENT_XPATH_STATUS
            )
            objStatus.click()

            # wait for the javascript to kick in
            time.sleep(1)

            # find open item in the checkbox
            objOpen = objDriver.find_element(
                'id',
                g.STR_ELEMENT_ID_CHECKBOX_OPEN
            )

            # check its status and click it if not ticked
            if objOpen.get_attribute(g.STR_ATTRIBUTE_CHECKED) is None:
                objOpen.click()

            # find approved item in the checkbox
            objApproved = objDriver.find_element(
                'id',
                g.STR_ELEMENT_ID_CHECKBOX_APPROVED
            )

            # check its status and click it if not ticked
            if objApproved.get_attribute(g.STR_ATTRIBUTE_CHECKED) is None:
                objApproved.click()

            # find closed item in the checkbox
            objClosed = objDriver.find_element(
                'id',
                g.STR_ELEMENT_ID_CHECKBOX_CLOSED
            )

            # check its status and click it if not ticked
            if objClosed.get_attribute(g.STR_ATTRIBUTE_CHECKED) is None:
                objClosed.click()

            # close the status dropdown
            objStatus.click()

            # set absence type
            objType = objDriver.find_element(
                'xpath',
                g.STR_ELEMENT_XPATH_TYPE
            )
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

            try:
                # attempt to locate the table
                objTable = objDriver.find_element(
                    'xpath',
                    g.STR_ELEMENT_XPATH_TABLE
                )

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
            except:
                # table not found, no absence available for this period
                lstAbsences = []

    return lstAbsences