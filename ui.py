# %%
# Contains methods and functions to handle user interaction to obtain calendar 
# data and submit absences to Xperience tracking system

# %% import modules
import globals as g
import calendar_handler as c
import website_handler as w

# %% define user interaction methods

def intGreeting():
    # set initial return value
    intContinue = -1

    # display the name and greet
    print(g.STR_UI_BOT_NAME)
    print(g.STR_UI_INTRO)

    # set a default value before looping
    strProcess = None

    # ask for an input until a valid answer is provided
    while strProcess not in g.LST_UI_ANSWERS:
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