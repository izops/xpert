# %% import modules
import re
import datetime
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define functions and methods that serve to obtain inputs from user for
# absence submission
def strGetCalendarConvention(pstrName):
    """Request user input for a valid calendar convention defined in a list.

    Inputs:
        - pstrName - string containing name to use in addressing the user

    Outputs:
        - strConvention - string value representing one of the possible choices
        of the calendar convention
    """
    # set the default value
    strConvention = ''

    # process the name
    if len(pstrName) > 0:
        strName = ', ' + pstrName
    else:
        strName = ''

    # ask user for the convention until valid answer is provided
    while strConvention not in g.LST_UI_ANSWERS_CONVENTION:
        strConvention = input(g.STR_UI_REQUEST_CONVENTION % strName)

    return strConvention

def strGetDate(pstrPlaceholderReplacement):
    """Request user input for a valid date in YYYYMMDD format.

    Inputs:
        - pstrPlaceholderReplacement - string for replacing the placeholder in
        the challenge message, usually 'start' or 'end'

    Outputs:
        - strDate - user input date in YYYYMMDD format
    """
    # ask user for a date
    strMessage = g.STR_UI_REQUEST_DATE.replace(
        g.STR_UI_REQUEST_PLACEHOLDER,
        pstrPlaceholderReplacement
    )

    # initialize user input variable
    strDate = ''

    # ask user for a date until a valid date is provided
    while not re.match(g.STR_REGEX_DATE, strDate):
        strDate = input(strMessage)

    # log obtained value
    logging.debug('strGetDate - strDate: ' + strDate)

    # save the match to an object
    objMatch = re.match(g.STR_REGEX_DATE, strDate)

    # extract parts of the date
    strYear, strMonth, strDay = objMatch.groups()
    
    # get today's date
    dttToday = datetime.date.today()

    # based on the found date parts, construct the input date
    if strYear is None:
        # missing year, use current
        strYear = str(dttToday.year)

    if strMonth is None:
        # missing month, use current and fix it to MM
        strMonth = str(dttToday.month + 100)
        strMonth = strMonth[1:]

    # log final parts of the date
    logging.debug('strGetDate - strYear: ' + strYear)
    logging.debug('strGetDate - strMonth: ' + strMonth)
    logging.debug('strGetDate - strDay: ' + strDay)

    # compile the final date to return
    strDate = strYear + strMonth + strDay

    return strDate