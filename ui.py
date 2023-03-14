# %%
# Contains methods and functions to handle user interaction to obtain calendar 
# data and submit absences to Xperience tracking system

# %% import modules
import globals as g
import calendar_handler as c
import website_handler as w

# %% define user interaction methods

def blnGreeting():
    # display the name
    print(g.STR_UI_BOT_NAME)

    # display the greeting and the options
    strProcess = input(g.STR_UI_INTRO)