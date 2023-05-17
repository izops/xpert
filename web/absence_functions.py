# %% import modules
import time
from selenium.webdriver.common.keys import Keys
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import general.general_functions as ggf

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define functions and methods that handle submission of absences
def blnOpenNewAbsence(pobjDriver, pstrAbsenceType):
    """Use selenium to open a new instance of absence submission in xperience.
    Check if the instance was opened successfully.

    Inputs:
        - pobjDriver - selenium webbrowser driver used to navigate the website
        - pstrAbsenceType - name of the absence instance that will be opened

    Outputs:
        - blnLoaded - boolean indicator of successful creating of a new absence
        instance
    """
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
    """Use selenium to fill details about specific absence. Check if the
    absence was submitted. If an error is identified, return it in a string.

    Inputs:
        - pobjDriver - selenium webdriver used for navigating websites
        - pstrAbsenceType - xperience absence type that is being filled in
        - pstrDateFrom - start date of the absence in DD/MM/YYYY format
        - pstrDateTo - optional argument, end date of the absence in DD/MM/YYYY
        format

    Outputs:
        - strError - error message retrieved from the website if any found
    """
    # convert the date to Slovak standard (dot separator)
    strDateFrom = pstrDateFrom.replace('/', '.')

    # log the date
    logging.debug('strAbsenceDetails - strDateFrom: ' + strDateFrom)

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

        # normalize non ascii characters
        strError = ggf.strNormalizeToASCII(strError)

        # replace all line breaks with a comma
        strError = strError.replace('\n', ', ')

        # add indentation and line breaks to the error message
        strError = '\t' + strError + '\n'

    except:
        # no error found
        strError = ''

    return strError