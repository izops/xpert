# %% import modules and scripts
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define UI functions for all processes
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

    if strProcess in g.LST_UI_ANSWERS_PROCESS[:g.INT_UI_CHOICES_MAX]:
        # the user accepted, change the indicator
        intContinue = int(strProcess)
    else:
        # user cancelled the process, greet and exit
        print(g.STR_UI_GOODBYE_CANCEL)

    return intContinue