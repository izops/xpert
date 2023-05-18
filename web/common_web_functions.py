# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define methods and functions that are shared accross features
def blnLogin(pobjDriver, pstrUserName, pstrPassword):
    """Use selenium webdriver to open xperience website and login using 
    the provided credentials. Check if the login was successfull.

    Inputs:
        - pobjDriver - selenium webbrowser driver used to navigate websites
        - pstrUserName - username for logging into xperience platform
        - pstrPassword - password for the website login page

    Outputs:
        - blnLoginSuccess - boolean indicator of success of the login process
    """
    # open the login url
    pobjDriver.get(g.STR_URL_LOGIN)

    # find the user id and password input fields, and login button
    objUserName = pobjDriver.find_element('id', g.STR_ELEMENT_ID_USERNAME)
    objPassword = pobjDriver.find_element('id', g.STR_ELEMENT_ID_PASSWORD)
    objLoginButton = pobjDriver.find_element('id', g.STR_ELEMENT_ID_LOGIN)

    # input username, password and click the login button
    objUserName.send_keys(pstrUserName)
    objPassword.send_keys(pstrPassword)
    objLoginButton.click()

    # attempt to find error box
    try:
        # the error box appeared, the login failed
        pobjDriver.find_element('id', g.STR_ELEMENT_ID_LOGIN_ERROR)

        # change the login indicator
        blnLoginSuccess = False
    except:
        # error box not found, login successful
        blnLoginSuccess = True

    return blnLoginSuccess