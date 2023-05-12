# %% import modules
import pandas as pd
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import web.scrape_absences as wsa

# %% define data processing functions
def dtfProcessDownloadedData(plstScrapedData):
    """Process the input list as a pandas data frame, convert data types, 
    rename columns and drop redundant fields.
    
    Inputs:
        - plstScrapedData - list of lists, each inner list representing single
        web-scraped data row

    Outputs:
        - dtfAbsences - pandas data frame containing 
    """
    # convert the inputs to data frame
    dtfAbsences = pd.DataFrame(plstScrapedData)

    # set all empty values and space values to missing
    dtfAbsences.replace('', pd.NA, inplace = True)
    dtfAbsences.replace(' ', pd.NA, inplace = True)

    # drop the empty columns
    dtfAbsences.dropna(axis = 1, how = 'all', inplace = True)

    # rename table columns
    dtfAbsences.columns = g.LST_COLUMN_NAMES_SCRAPED

    # replace date separators with slash, decimal comma with a dot
    dtfAbsences = dtfAbsences.apply(
        lambda x: x.replace(
            {'\.': '/', ',': '.'},
            regex=True
        )
    )

    # extract day info from the duration column
    dtfAbsences['Duration'] = dtfAbsences['Duration'].str.extract(
        g.STR_REGEX_ABSENCE_DURATION
    )

    # convert date columns to date type
    dtfAbsences[['From', 'To']] = dtfAbsences[['From', 'To']].apply(
        pd.to_datetime,
        dayfirst = True
    )

    # convert number column to numeric
    dtfAbsences['Duration'] = pd.to_numeric(dtfAbsences['Duration'])

    # drop non necessary columns
    dtfAbsences.drop(columns = ['Name', 'Modified'], inplace = True)

    return dtfAbsences

def ObtainXperienceAbsences(
    pstrUserName,
    pstrPassword,
    pstrDateFrom,
    pstrDateTo,
    pstrAbsenceType
):
    """Coordinate process of scraping web data from xperience, cleaning the
    data and outputting them to an external file.

    Inputs:
        - pstrUserName - username for xperience login
        - pstrPassword - password for xperience login
        - pstrDateFrom - start date of absence filtering, YYYYMMDD string
        - pstrDateTo - end date of absence filtering, YYYYMMDD string
        - pstrAbsenceType - absence type to scrape

    Outputs:
        - None returned, external txt file with scraped data created
    """
    # scrape data from web based on user request
    lstScrapedAbsences = wsa.lstDownloadData(
        pstrUserName,
        pstrPassword,
        pstrDateFrom,
        pstrDateTo,
        pstrAbsenceType
    )

    # process the absences
    dtfProcessedAbsences = dtfProcessDownloadedData(lstScrapedAbsences)

    # save the processed data to an external file
    dtfProcessedAbsences.to_csv(g.STR_FULL_PATH_SCRAPED_DATA, sep = '\t')