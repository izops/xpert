# %% import modules
import pandas as pd
import os
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% functions and methods to import absences from text input to calendar
def blnVerifyData(pdtfToVerify):
    # set a default value to return
    blnValidData = True
    
    # verify the number of columns
    if pdtfToVerify.shape[1] != len(g.LST_COLUMN_NAMES_IMPORT):
        blnValidData = False

    # verify the names of columns
    if blnValidData \
    and [
        x.capitalize() for x in pdtfToVerify.columns
    ] != g.LST_COLUMN_NAMES_IMPORT:
        blnValidData = False

    # verify the number of rows
    if blnValidData and pdtfToVerify.shape[0] < 1:
        blnValidData = False

    return blnValidData

def ImportAbsences():
    if os.path.isfile(g.STR_FULL_PATH_SCRAPED_DATA):
        # read the data to import
        dtfSource = pd.read_csv(g.STR_FULL_PATH_SCRAPED_DATA, sep = '\t')

        # check if the data is relevant for importing
        if blnVerifyData(dtfSource):
            # data is relevant, import it
            pass
        else:
            # data not relevant, inform the user
            strMessage = 'Nothing relevant to import in the source data'
            print(strMessage)

    else:
        # if import failed, inform the user
        strMessage = 'Import of absences to Outlook failed, unable to import '
        strMessage += os.path.abspath(g.STR_FULL_PATH_SCRAPED_DATA)
        print(strMessage)