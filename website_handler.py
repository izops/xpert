# %% import modules
# selenium elements
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# other modules
import getpass
import time

# import scripts
import globals as g

# %% define user interaction functions
def strGetUserName():
    # get user name from the local environment
    strUserName = getpass.getuser() + g.STR_USER_DOMAIN

    return strUserName

def strGetPassword():
    # put together prompt message
    strMessage = 'Please, provide your password.'
    strMessage += ' (it will be hidden and discarded afterwards)\n'

    # ask user for their password
    strPassword = getpass.getpass(strMessage)

    return strPassword

# %% define functions for handling the browser
def objRunProcess():
    # get user name and password
    strUserName = strGetUserName()
    strPassword = strGetPassword()

    # set up the webdriver
    objDriver = webdriver.Edge()

    # log in to the page
    blnContinue = blnLogin(objDriver, strUserName, strPassword)

    # discard the password
    del strPassword


def blnLogin(pobjDriver, pstrUserName, pstrPassword):
    # open the login url
    pobjDriver.get(g.STR_URL_LOGIN)

    # find the user id and password input fields, and login button
    objUserName = pobjDriver.find_element('id', g.STR_ELEMENT_ID_USERNAME)
    objPassword = objDriver.find_element('id', g.STR_ELEMENT_ID_PASSWORD)
    objLoginButton = pobjDriver.find_element('id', g.STR_ELEMENT_ID_LOGIN)

    # open the login page
    objDriver.get(g.STR_URL_LOGIN)

    # input username, password and click the login button
    objUserName.send_keys(pstrUserName)
    objPassword.send_keys(pstrPassword)
    objLoginButton.click()

    # attempt to find error box
    try:
        # the error box appeared, the login failed
        pobjDriver.find_element('id', g.STR_ELEMENT_ID_LOGIN_ERROR)

        # change the login indicator
        blnLoginSuccess = False
    except:
        # error box not found, login successful
        blnLoginSuccess = True

    return blnLoginSuccess

def blnOpenNewAbsence(pobjDriver, pstrAbsenceType):
    # identify element with the absence type dropdown
    objAbsenceType = pobjDriver.find_element(
        'xpath',
        g.STR_ELEMENT_XPATH_ABSENCE_TYPE
    )

    # select the absence type
    objAbsenceType.send_keys(pstrAbsenceType)

    # wait for javascript to load the dropdown menu
    time.sleep(1)

    # select the absence type from the dropdown and confirm
    objAbsenceType.send_keys(Keys.TAB)
    objAbsenceType.send_keys(Keys.ENTER)

    # check if the absence detail loaded or not
    try:
        pobjDriver.find_element('xpath', g.STR_ELEMENT_XPATH_ABSENCE_DETAIL)

        # absence detail found on the page, set the indicator to positive
        blnLoaded = True
    except:
        # the absence detail not loaded, change the indicator to negative
        blnLoaded = False

    return blnLoaded

def blnAbsenceDetails(
    pobjDriver,
    pstrAbsenceType,
    pstrDateFrom,
    pstrDateTo = ''
):
    # set up the interaction based on the absence type
    if pstrAbsenceType == g.STR_ABSENCE_TYPE_HOME_OFFICE:
        # find the start date input field
        objDateStart = pobjDriver.find_element(
            'xpath',
            g.STR_ELEMENT_XPATH_ABSENCE_DATE_START
        )

        # type in the start date
        objDateStart.send_keys(pstrDateFrom)

        # find the end date if applicable
        if len(pstrDateTo) > 0 and pstrDateFrom != pstrDateTo:
            objDateEnd = pobjDriver.find_element(
                'xpath',
                g.STR_ELEMENT_XPATH_ABSENCE_DATE_END
            )

            # type in the end date
            objDateEnd.send_keys(pstrDateTo)

    # for safety reasons activate notes field
    objNotes = pobjDriver.find_element(
        'id',
        g.STR_ELEMENT_ID_ABSENCE_NOTE
    )
    objNotes.click()

    # localize the submit button
    objSubmit = pobjDriver.find_element(
        'id',
        g.STR_ELEMENT_ID_ABSENCE_SUBMIT
    )

    # submit the absence
    objSubmit.click()

    # verify that the page loaded back to the main menu
    try:
        pobjDriver.find_element(
            'xpath',
            g.STR_ELEMENT_XPATH_ABSENCE_MAIN
        )

        # main menu title found, the absence was submitted successfully
        blnSubmitted = True
    except:
        # main menu title not found, absence failed
        blnSubmitted = False

    return blnSubmitted


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