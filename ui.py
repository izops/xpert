# %%
# Contains methods and functions to handle user interaction to obtain calendar 
# data and submit absences to Xperience tracking system

# %% import modules
import re

import globals as g
import calendar_handler as c
import website_handler as w

# %% define user interaction methods
def intGreeting():
    '''
    Outputs the introduction and asks for the type of work that should be done

    Inputs:
        - None

    Outputs:
        - intContinue - indication of the process that should be run
    '''
    # set initial return value
    intContinue = -1

    # display the name and greet
    print(g.STR_UI_BOT_NAME)
    print(g.STR_UI_INTRO)

    # set a default value before looping
    strProcess = None

    # ask for an input until a valid answer is provided
    while str(strProcess).lower() not in g.LST_UI_ANSWERS:
        # add a warning if this is not the first time
        if not strProcess is None:
            print('Please, input values from the provided list')

        # display the options
        strProcess = input(g.STR_UI_OFFER)

    if strProcess in g.LST_UI_ANSWERS[:3]:
        # the user accepted, change the indicator
        intContinue = int(strProcess)
    else:
        # user cancelled the process, greet and exit
        print(g.STR_UI_GOODBYE_CANCEL)

    return intContinue

def strGetDate(pstrPlaceholderReplacement):
    '''
    Requests user input for a valid date in YYYYMMDD format

    Inputs:
        - pstrPlaceholderReplacement - string for replacing the placeholder in
        the challenge message, usually 'start' or 'end'

    Outputs:
        - strDate - user input date in YYYYMMDD format
    '''
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

def RunProcess(pintChoice):
    # analyze the calendar
    if pintChoice in [g.INT_UI_CHOICE_CALENDAR, g.INT_UI_CHOICE_ALL]:
        # request the starting and ending point for the calendar analysis
        strDateStart = strGetDate(g.STR_UI_REQUEST_DATE_START)
        strDateEnd = strGetDate(g.STR_UI_REQUEST_DATE_END)

        # launch the calendar analysis
        c.AnalyzeCalendar(strDateStart, strDateEnd, True)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

        if pintChoice == g.INT_UI_CHOICE_CALENDAR:
            # the process ends here, part with the user
            print(g.STR_UI_GOODBYE)
        
    if pintChoice in [g.INT_UI_CHOICE_XPERIENCE, g.INT_UI_CHOICE_ALL]:
        # run xperience process
        pass