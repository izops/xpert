# %% import modules
import unicodedata
import datetime
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

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

def strChangeDateFormat(pstrDateIn, pstrDelimiter, pblnToDMY = True):
    """Change form of the date from YYYYMMDD to DDMMYYYY or vice versa using
    selected delimiter, independently of input date delimiter.

    Inputs:
        - pstrDateIn - string date to change
        - pstrDelimiter - string delimiter to be used in output date

    Output:
        - strDateOut - date in the new format
    
    """
    if pblnToDMY:
        # remove delimiters from YYYYMMDD 
        strDateIn = pstrDateIn.replace(' ', '')
        strDateIn = strDateIn.replace('.', '')
        strDateIn = strDateIn.replace('/', '')
        strDateIn = strDateIn.replace('-', '')

        # extract date parts of YYYYMMDD
        strYear = strDateIn[:4]
        strMonth = strDateIn[4:6]
        strDay = strDateIn[6:]

        # add to a list
        lstDate = [strDay, strMonth, strYear]

    else:
        # remove delimiters from DDMMYYYY 
        strDateIn = pstrDateIn.replace(' ', '')
        strDateIn = strDateIn.replace('.', '')
        strDateIn = strDateIn.replace('/', '')
        strDateIn = strDateIn.replace('-', '')

        # extract date parts of DDMMYYYY
        strDay = strDateIn[:2]
        strMonth = strDateIn[2:4]
        strYear = strDateIn[4:]

        # add to a list
        lstDate = [strYear, strMonth, strDay]

    # change the date to the requested format with specified delimiter
    strDateOut = pstrDelimiter.join(lstDate)

    return strDateOut

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