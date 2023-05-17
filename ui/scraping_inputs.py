# %% import modules
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

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

    # prepare lowercase list of values to compare against
    lstAbsences = [x.lower() for x in g.LST_SCRAPE_ABSENCES]

    # prepare list of indexes for alternative choice
    lstIndexes = [str(x + 1) for x in range(len(lstAbsences))]

    # obtain from user valid answer, make difference between first
    # and the rest of the inputs
    while strResponse is None \
    or (
        strResponse != 'c' \
        and not strResponse in lstAbsences \
        and not strResponse in lstIndexes
    ):
        # log the response
        logging.debug('strGetAbsenceType - strResponse: ' + str(strResponse))

        # loop until either canceled, or valid absence is provided
        if strResponse is None:
            # user is prompted for input for the first time
            strMessage = 'Please, select an absence to scrape from the '
            strMessage += 'following list: \n'
            strMessage += str(g.LST_SCRAPE_ABSENCES) + '\n'
            strMessage += 'or type in its respective index (for example '
            strMessage += g.LST_SCRAPE_ABSENCES[0] + ' = ' + lstIndexes[0]
            strMessage += '):\n'
        
        elif strResponse != 'c' \
        or not strResponse in lstAbsences \
        or not strResponse in lstIndexes:
            # not the first input but not valid either
            strMessage = 'Apologies, I cannot scrape "' + strResponse 
            strMessage += '" absence for you. '
            strMessage += 'Choose a different one or type "c" to cancel.\n'
            
        # prompt user for input, make it lowercase
        strResponse = input(strMessage).lower()

    # if the input was an index, convert it to an absence value
    if strResponse.isnumeric():
        # subtract one to get zero-based index for python
        strResponse = g.LST_SCRAPE_ABSENCES[int(strResponse) - 1]

    return strResponse