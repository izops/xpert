# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define functions for obtaining inputs for data scraping
def strGetAbsenceType():
    """Obtain valid absence type for scraping from user.
    
    Inputs:
        - None

    Outputs:
        - strResponse - string containing user response to absence prompt
    """
    # set initial value
    strResponse = None

    # obtain from user valid answer, make difference between first
    # and the rest of the inputs
    while strResponse is None \
    or (
        strResponse != 'c' \
        and not strResponse in g.LST_SCRAPE_ABSENCES
    ):
        # loop until either canceled, or valid absence is provided
        if strResponse is None:
            # user is prompted for input for the first time
            strMessage = 'Please, select an absence to scrape from the '
            strMessage += 'following list: \n'
            strMessage += str(g.LST_SCRAPE_ABSENCES)
        
        elif strResponse != 'c' \
        or not strResponse in g.LST_SCRAPE_ABSENCES:
            # not the first input but not valid either
            strMessage = 'Apologies, I cannot scrape "' + strResponse 
            strMessage += '" absence for you. '
            strMessage += 'Choose a different one or type "c" to cancel.\n'
            
        # prompt user for input, make it lowercase
        strResponse = input(strMessage).lower()

    return strResponse