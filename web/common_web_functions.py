# %% import modules
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define methods and functions that are shared accross features
def blnLogin(pobjDriver: webdriver):
    """Use selenium webdriver to open xperience website and login using 
    single sign on. Check if the login was successfull.

    Inputs:
        - pobjDriver - selenium webbrowser driver used to navigate websites

    Outputs:
        - blnLoginSuccess - boolean indicator of success of the login process
    """
    # open the login url
    pobjDriver.get(g.STR_URL_LOGIN)

    try:
        # wait for the single sign on to finish loading
        objHomeButton = WebDriverWait(
            pobjDriver,
            g.FLT_MAX_LOGIN_TIMEOUT
        ).until(
            EC.presence_of_element_located(('id', g.STR_ELEMENT_ID_HOME))
        )

        # single sign on and the subsequent page loaded successfully
        blnLoginSuccess = True
    except:
        # single sign on didn't work
        blnLoginSuccess = False

    return blnLoginSuccess