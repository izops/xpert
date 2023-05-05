# %%
# Contains functions and methods that read data inputs for Xperience submission,
# and that operate the Xperience website and submit the absence details. 
# Currently submits only full-day home office absences.

# %% import modules
# selenium elements
from selenium import webdriver

# other modules
import re
import os
import sys

# import scripts
sys.path.append('../emea_oth_xpert/')
import general.global_constants as g
import web.common_web_fucntions as cwf
import web.absence_functions as was
import general.general_functions as gf

# %% define functions for handling the browser
def lstReadData(pstrPath):
    """Read calendar data saved in a text file.

    Inputs:
        - pstrPath - full path to the file with the calendar data saved

    Outputs:
        - lstCalendarData - list of the parsed calendar data containing tuples
        of start date, end date and absence type
    """
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

# %% define process handling
def SubmitAbsences(pstrUserName, pstrPassword):
    """Handle entire process of obtaining login info, logging into xperience
    website, reading calendar data and submitting all relevant entries.

    Inputs:
        - pstrUserName - string containing username for xperience login
        - pstrPassword - string containing password for xperience login

    Outputs:
        - None, but the entire process is run
    """
    # read the calendar data from the source file
    try:
        lstCalendarData = lstReadData(g.STR_FULL_PATH_CALENDAR_DATA)
    except:
        print('The file doesn\'t exist.')

    # continue if there is any calendar entry
    if len(lstCalendarData) > 0:
        # initialize a message variable
        strMessage = ''

        # set up the webdriver
        objDriver = webdriver.Edge()

        # log in to the page
        blnContinue = cwf.blnLogin(objDriver, pstrUserName, pstrPassword)

        # discard the password
        del pstrPassword

        if blnContinue:
            # if login successful, add the absence entry for each data point
            for tplAbsence in lstCalendarData:
                # work with home office data for now
                if tplAbsence[
                    2
                ].lower() == g.STR_ABSENCE_TYPE_HOME_OFFICE.lower():
                    # open a new absence entry
                    blnContinue = was.blnOpenNewAbsence(
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
                    strError = was.strAbsenceDetails(
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
        gf.WriteLog(strMessage)

    # print the results of the process
    print(strMessage)