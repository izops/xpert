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
    while str(strProcess).lower() not in g.LST_UI_ANSWERS_PROCESS:
        # add a warning if this is not the first time
        if not strProcess is None:
            print('Please, input values from the provided list')

        # display the options
        strProcess = input(g.STR_UI_OFFER)

    if strProcess in g.LST_UI_ANSWERS_PROCESS[:3]:
        # the user accepted, change the indicator
        intContinue = int(strProcess)
    else:
        # user cancelled the process, greet and exit
        print(g.STR_UI_GOODBYE_CANCEL)

    return intContinue

def strGetCalendarConvention():
    '''
    Requests user input for a valid calendar convention defined in a list

    Inputs:
        - None

    Outputs:
        - strConvention - string value representing one of the possible choices
        of the calendar convention
    '''
    # set the default value
    strConvention = ''

    # ask user for the convention until valid answer is provided
    while strConvention not in g.LST_UI_ANSWERS_CONVENTION:
        strConvention = input(g.STR_UI_REQUEST_CONVENTION)

    return strConvention

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

# %% define the master method to launch the process parts
def RunProcess(pintChoice):
    '''
    Based on the input runs Outlook calendar analysis, submission of absences to
    Xperience or both

    Inputs:
        - pintChoice - numeric indication of the process to be run

    Outputs:
        - None, either one or two processes are run
    '''
    # analyze the calendar
    if pintChoice in [g.INT_UI_CHOICE_CALENDAR, g.INT_UI_CHOICE_ALL]:
        # request the starting and ending point for the calendar analysis
        strDateStart = strGetDate(g.STR_UI_REQUEST_DATE_START)
        strDateEnd = strGetDate(g.STR_UI_REQUEST_DATE_END)

        # get the calendar convention
        strConvention = strGetCalendarConvention()

        # create boolean parameter based on the output from convention check
        if strConvention == g.LST_UI_ANSWERS_CONVENTION[0]:
            blnOfficeFocused = True
        else:
            blnOfficeFocused = False

        # launch the calendar analysis
        c.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

        if pintChoice == g.INT_UI_CHOICE_CALENDAR:
            # the process ends here, say goodbye to the user
            print(g.STR_UI_GOODBYE)
        
    if pintChoice in [g.INT_UI_CHOICE_XPERIENCE, g.INT_UI_CHOICE_ALL]:
        # run xperience process
        w.SubmitAbsences()

        # the process ends here, say goodbye to the user
        print(g.STR_UI_GOODBYE)