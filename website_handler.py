# %%
# Contains functions and methods that read data inputs for Xperience submission,
# and that operate the Xperience website and submit the absence details. 
# Currently submits only full-day home office absences.

# %% import modules
# selenium elements
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# other modules
import getpass
import time
import re
import os

# import scripts
import globals as g

# %% define user interaction functions
def strGetUserName():
    '''
    Looks into OS environment and retrieves the username of the current user,
    appends the name of the domain to create a corporate email address

    Inputs:
        - None

    Outputs:
        - strUserName - corporate email address based on environment username
    '''
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
def blnLogin(pobjDriver, pstrUserName, pstrPassword):
    '''
    Uses selenium webdriver to open xperience website and login using 
    the provided credentials. Checks if the login was successfull.

    Inputs:
        - pobjDriver - selenium webbrowser driver used to navigate websites
        - pstrUserName - username for logging into xperience platform
        - pstrPassword - password for the website login page

    Outputs:
        - blnLoginSuccess - boolean indicator of success of the login process
    '''
    # open the login url
    pobjDriver.get(g.STR_URL_LOGIN)

    # find the user id and password input fields, and login button
    objUserName = pobjDriver.find_element('id', g.STR_ELEMENT_ID_USERNAME)
    objPassword = pobjDriver.find_element('id', g.STR_ELEMENT_ID_PASSWORD)
    objLoginButton = pobjDriver.find_element('id', g.STR_ELEMENT_ID_LOGIN)

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
    '''
    Uses selenium to open a new instance of absence submission in xperience.
    Checks if the instance was opened successfully.

    Inputs:
        - pobjDriver - selenium webbrowser driver used to navigate the website
        - pstrAbsenceType - name of the absence instance that will be opened

    Outputs:
        - blnLoaded - boolean indicator of successful creating of a new absence
        instance
    '''
    # open the URL for adding a new absence
    pobjDriver.get(g.STR_URL_ADD_ABSENCE)

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

def strAbsenceDetails(
    pobjDriver,
    pstrAbsenceType,
    pstrDateFrom,
    pstrDateTo = ''
):
    '''
    Uses selenium to fill details about specific absence. Checks if the absence
    was submitted. If an error is identified, returns it in a string.

    Inputs:
        - pobjDriver - selenium webdriver used for navigating websites
        - pstrAbsenceType - xperience absence type that is being filled in
        - pstrDateFrom - start date of the absence in DD/MM/YYYY format
        - pstrDateTo - optional argument, end date of the absence in DD/MM/YYYY
        format

    Outputs:
        - strError - error message retrieved from the website if any found
    '''
    # convert the date to Slovak standard (dot separator)
    strDateFrom = pstrDateFrom.replace('/', '.')

    # set up the interaction based on the absence type
    if pstrAbsenceType == g.STR_ABSENCE_TYPE_HOME_OFFICE:
        # find the start date input field
        objDateStart = pobjDriver.find_element(
            'xpath',
            g.STR_ELEMENT_XPATH_ABSENCE_DATE_START
        )

        # type in the start date
        objDateStart.send_keys(strDateFrom)

        # find the end date if applicable
        if len(pstrDateTo) > 0 and pstrDateFrom != pstrDateTo:
            # convert the date to Slovak standard (dot separator)
            strDateTo = pstrDateTo.replace('/', '.')

            # locate end date field
            objDateEnd = pobjDriver.find_element(
                'xpath',
                g.STR_ELEMENT_XPATH_ABSENCE_DATE_END
            )

            # activate the field to activate autofill by sending null key
            objDateEnd.send_keys(Keys.NULL)

            # clear the contents of the input field
            objDateEnd.clear()

            # type in the end date
            objDateEnd.send_keys(strDateTo)

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
        # attempt to locate error box
        objError = pobjDriver.find_element(
            'xpath',
            g.STR_ELEMENT_XPATH_ABSENCE_ERROR
        )

        # read the error message
        strError = objError.text

        # replace all line breaks with a comma
        strError = strError.replace('\n', ', ')

        # add indentation and line breaks to the error message
        strError = '\t' + strError + '\n'

    except:
        # no error found
        strError = ''

    return strError

def lstReadData(pstrPath):
    '''
    Reads calendar data saved in a text file.

    Inputs:
        - pstrPath - full path to the file with the calendar data saved

    Outputs:
        - lstCalendarData - list of the parsed calendar data containing tuples
        of start date, end date and absence type
    '''
    # verify the file exists
    assert os.path.isfile(pstrPath), 'The file ' + pstrPath + ' doesn\'t exist'

    # open the file, read only
    objCalendarData = open(pstrPath, 'r')

    # initialize a list of values read from the source file
    lstCalendarData = []

    # read in the file
    for strRow in objCalendarData:
        # try to match the row with regex
        objMatch = re.match(g.STR_REGEX_DATA_INPUT, strRow, re.IGNORECASE)

        if objMatch:
            # there is a match, parse the data into a tuple
            tplEntry = objMatch.group(1), objMatch.group(2), objMatch.group(3)

            # add the data point to the output list
            lstCalendarData.append(tplEntry)

    # close the file
    objCalendarData.close()

    return lstCalendarData

def WriteLog(pstrLogText):
    '''
    Opens a text file in the data folder and saves contents of a string to it.

    Inputs:
        - pstrLogText - string containing text data to be saved in the log

    Outputs:
        - none, log file is created in the data folder
    '''
    # open a new file
    objLog = open(g.STR_FULL_PATH_LOG, 'w')

    # write the log contents
    objLog.write(pstrLogText)

    # close the file
    objLog.close()

# %% define process handling
def SubmitAbsences():
    '''
    Handles entire process of obtaining login info, logging into xperience
    website, reading calendar data and submitting all relevant entries

    Inputs:
        - None

    Outputs:
        - None, but the entire process is run
    '''
    # read the calendar data from the source file
    try:
        lstCalendarData = lstReadData(g.STR_FULL_PATH_CALENDAR_DATA)
    except:
        print('The file doesn\'t exist.')

    # continue if there is any calendar entry
    if len(lstCalendarData) > 0:
        # initialize a message variable
        strMessage = ''

        # get user name and password
        strUserName = strGetUserName()
        strPassword = strGetPassword()

        # set up the webdriver
        objDriver = webdriver.Edge()

        # log in to the page
        blnContinue = blnLogin(objDriver, strUserName, strPassword)

        # discard the password
        del strPassword

        if blnContinue:
            # if login successful, add the absence entry for each data point
            for tplAbsence in lstCalendarData:
                # work with home office data for now
                if tplAbsence[
                    2
                ].lower() == g.STR_ABSENCE_TYPE_HOME_OFFICE.lower():
                    # open a new absence entry
                    blnContinue = blnOpenNewAbsence(
                        objDriver,
                        tplAbsence[2].lower()
                    )

                    # use the end date only if different from the start date
                    if tplAbsence[0] != tplAbsence[1]:
                        strEndDate = tplAbsence[1]
                    else:
                        # don't use an end date
                        strEndDate = ''

                    # submit the absence
                    strError = strAbsenceDetails(
                        objDriver,
                        tplAbsence[2],
                        tplAbsence[0],
                        strEndDate
                    )

                    # prepare an error message in case the process failed
                    if len(strError) > 0:
                        # submission check not passed, prepare error message
                        strMessage += 'FAIL: Submission of absence from '
                        strMessage += tplAbsence[0] + ' to ' + tplAbsence[1]
                        strMessage += ', ' + tplAbsence[2] + ', failed with '
                        strMessage += 'the following error:\n'
                        strMessage += strError
                    else:
                        # submission successfull, add the message to the log
                        strMessage += 'SUCCESS: Submission of absence from '
                        strMessage += tplAbsence[0] + ' to ' + tplAbsence[1]
                        strMessage += ', ' + tplAbsence[2] + ', succeeded\n'

                else:
                    # unsupported type of absence
                    strMessage = 'Unsupported type of absence: ' + tplAbsence[2]

        else:
            # login failed, prepare the message
            strMessage = 'Login failed. Either you provided incorrect'
            strMessage += ' credentials or your password expired.'

    else:
        # there are no calendar entries to submit
        strMessage = 'There were 0 entries read from the data file, '
        strMessage += 'the process ends here.'

    # save the log as an external file if any absence was submitted
    if strMessage.find('\n') >= 0:
        WriteLog(strMessage)

    # print the results of the process
    print(strMessage)