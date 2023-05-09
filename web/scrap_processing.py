# %% import modules
import pandas as pd
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% define data processing functions
def dtfProcessDownloadedData(plstScrapedData):
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
    dtfAbsences.duration = dtfAbsences.duration.str.extract(
        g.STR_REGEX_ABSENCE_DURATION
    )

    # convert date columns to date type
    dtfAbsences[['From', 'To']] = dtfAbsences[['From', 'To']].apply(
        pd.to_datetime,
        dayfirst = True
    )

    # convert number column to numeric
    dtfAbsences.duration = pd.to_numeric(dtfAbsences.duration)

    # drop non necessary columns
    dtfAbsences.drop(columns = ['Name', 'Modified'], inplace = True)

    return dtfAbsences