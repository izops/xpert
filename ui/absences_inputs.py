# %% import modules
import re
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define functions and methods that serve to obtain inputs from user for
# absence submission
def strGetCalendarConvention():
    """Request user input for a valid calendar convention defined in a list.

    Inputs:
        - None

    Outputs:
        - strConvention - string value representing one of the possible choices
        of the calendar convention
    """
    # set the default value
    strConvention = ''

    # ask user for the convention until valid answer is provided
    while strConvention not in g.LST_UI_ANSWERS_CONVENTION:
        strConvention = input(g.STR_UI_REQUEST_CONVENTION)

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

    return strDate