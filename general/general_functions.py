# %% import modules
import unicodedata
import datetime

import globals as g

# %% define general functions unrelated specifically to any process
def dttConvertDate(pstrYYYYMMDD):
    """Convert YYYYMMDD string date to datetime object.

    Inputs:
        - pstrYYYYMMDD - string date in YYYYMMDD format

    Outputs:
        - dttConverted - converted date as a datetime object
    """
    # parse the data from string and convert to datetime
    dttConverted = datetime.datetime(
        int(pstrYYYYMMDD[:4]),
        int(pstrYYYYMMDD[4:6]),
        int(pstrYYYYMMDD[6:])
    )

    return dttConverted

def strConvertDate(pdttDateTime):
    """Convert datetime object into date in YYYYMMDD format.

    Inputs:
        - pdttDateTime - datetime object

    Outputs:
        - strConverted - string representation of a date in YYYYMMDD format
    """
    # convert date to YYYYMMDD date
    strConverted = pdttDateTime.strftime('%Y%m%d')

    return strConverted

def strNormalizeToASCII(pstrInput):
    """Normalize non-ASCII string to the best corresponding ASCII match.

    Inputs:
        - pstrInput - input string to clean up

    Outputs:
        - string cleaned of all non ascii characters to their best match
    """
    # remove all non ASCII characters from the string
    strOutput = ''.join(
        strChar for strChar in unicodedata.normalize('NFD', pstrInput)
        if unicodedata.category(strChar) != 'Mn'
    )

    return strOutput

def WriteLog(pstrLogText):
    """Open a text file in the data folder and save contents of a string to it.

    Inputs:
        - pstrLogText - string containing text data to be saved in the log

    Outputs:
        - None, log file is created in the data folder
    """
    # open a new file
    objLog = open(g.STR_FULL_PATH_LOG, 'w')

    # write the log contents
    objLog.write(pstrLogText)

    # close the file
    objLog.close()