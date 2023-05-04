# %% import modules
import maskpass
import os
import globals as g

# %% define functions to obtain credentials
def strGetUserName():
    """Look into OS environment and retrieve the username of the current user,
    append the name of the domain to create a corporate email address.

    Inputs:
        - None

    Outputs:
        - strUserName - corporate email address based on environment username
    """
    # get user name from the local environment
    strUserName = os.getlogin() + g.STR_USER_DOMAIN

    return strUserName

def strGetPassword():
    """Prompt user to obtain password, hide it as a masked input and return it.

    Inputs:
        - None

    Outputs:
        - strPassword - string containing user input
    """
    # put together prompt message
    strMessage = 'Please, provide your password.'
    strMessage += ' (it will be hidden and discarded afterwards)\n'

    # ask user for their password
    strPassword = maskpass.askpass(strMessage)

    return strPassword