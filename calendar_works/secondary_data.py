import win32com.client as win32
import logging
import sys

# import scripts
sys.path.append('../emea_oth_xpert/')
import general.global_constants as g

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define functions and methods to obtain additional data from Outlook
def tplGetFullUserName():
    """Get user name and surname from Outlook.
    
    Inputs:
        - None

    Outputs:
        - tuple containing name and surname of currently logged in Outlook user
    """

    # set the application instance
    objApp = win32.Dispatch('Outlook.Application')

    # set the default name space
    objNameSpace = objApp.GetNamespace('MAPI')

    # get the current user
    objUser = objNameSpace.CurrentUser

    # obtain the full name
    strFullName = objUser.Name

    logging.debug

    # make it a tuple
    tplFullName = tuple(strFullName.split(' '))

    return tplFullName