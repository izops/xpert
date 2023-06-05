# %% import modules and scripts
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define UI functions for all processes
def intGreeting(ptplFullName):
    '''
    Outputs the introduction and asks for the type of work that should be done

    Inputs:
        - ptplFullName - tuple containing full name of the active user

    Outputs:
        - intContinue - indication of the process that should be run
    '''
    # set initial return value
    intContinue = -1

    # process the user name for intro
    if len(ptplFullName[0]) > 0:
        strName = ' ' + ptplFullName[0]
    else:
        strName = ''

    # display the name and greet
    print(g.STR_UI_BOT_NAME)
    print(g.STR_UI_INTRO % strName)

    # process the user name for the rest
    if len(ptplFullName[0]) > 0:
        strName = ', ' + ptplFullName[0]
    else:
        strName = ''

    # set a default value before looping
    strProcess = None

    # ask for an input until a valid answer is provided
    while str(strProcess).lower() not in g.LST_UI_ANSWERS_PROCESS:
        # log current input
        logging.debug('intGreeting - strProcess: ' + str(strProcess))

        # add a warning if this is not the first time
        if not strProcess is None:
            print('Please, input values from the provided list')

        # display the options
        strProcess = input(g.STR_UI_OFFER % strName)

    if strProcess in g.LST_UI_ANSWERS_PROCESS[:g.INT_UI_CHOICES_MAX]:
        # the user accepted, change the indicator
        intContinue = int(strProcess)
    else:
        # user cancelled the process, greet and exit
        print(g.STR_UI_GOODBYE_CANCEL % strName)

    # log return value
    logging.info('intGreeting - intContinue: ' + str(intContinue))

    return intContinue