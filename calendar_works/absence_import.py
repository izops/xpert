# %% import modules
import pandas as pd
import win32com.client as win32
import datetime
import os
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import general.general_functions as ggf

# %% functions and methods to import absences from text input to calendar
def blnVerifyData(pdtfToVerify):
    """Verify the data for import.

    Inputs:
        - pdtfToVerify - pandas dataframe that will be verified for validity

    Outpust:
        - blnValidData - boolean indicator if the file is valid for import to
        Outlook
    """
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

def lstSplitOffHalfDays(pdtfAbsenceData):
    # convert the absences to a list
    lstAbsences = pdtfAbsenceData.values.tolist()

    # split off half-day absences and append them to the end of the existing
    # chunk
    for intAbsence in range(len(lstAbsences)):
        # split the half day absence from a continuous full day absence
        if int(lstAbsences[intAbsence][2]) != lstAbsences[intAbsence][2] \
        and lstAbsences[intAbsence][2] > 1:
            # create a copy of the list
            lstHalfDay = lstAbsences[intAbsence].copy()

            # change the duration of the absence chunks
            lstAbsences[intAbsence][2] = float(int(lstAbsences[intAbsence][2]))
            lstHalfDay[2] = 0.5

            # change the start date of the second absence
            lstHalfDay[0] = lstHalfDay[1]

            # change the end date of the first absence
            dttEnd = ggf.dttConvertDate(
                lstAbsences[intAbsence][1],
                '-'
            ) - datetime.timedelta(days = 1)

            # assign the value of the first absence without the half day
            lstAbsences[intAbsence][1] = dttEnd.strftime('%Y-%m-%d')

            # insert the half day chunk to the list of absences
            lstAbsences.insert(intAbsence + 1, lstHalfDay)

    return lstAbsences

def intDetermineHalfDay(plstAbsence1, plstAbsence2):
    """Determine correct position of half day absence.
    
    Inputs:
        - plstAbsence1 - list containing all data for an absence
        - plstAbsence2 - list containing all data for an immediate next absence

    Outputs:
        - intHalfDay - integer flag denoting if the absence starts in first or
        second half of the day
    """
    # set the default return value to first half of day
    intHalfDay = 1

    # compute half day if second absence exists
    if not plstAbsence2 is None:
        # check if the half day starts a chunk of absences
        if plstAbsence1[2] == 0.5:
            # convert the dates required for checks
            # current end, next start
            dttHalf = ggf.dttConvertDate(
                plstAbsence1[1],
                '-'
            )
            dttNext = ggf.dttConvertDate(
                plstAbsence2[0],
                '-'
            )

            if dttHalf + datetime.timedelta(days = 1) == dttNext \
            and plstAbsence2[2] >= 1:
                # half day is followed by a full day absence, set it to
                # second half of the day
                intHalfDay = 2

    return intHalfDay

def SaveAbsences(plstAbsenceData):
    """Save all absences from input to Outlook calendar.
    
    Inputs:
        - plstAbsenceData - list of absences converted from pandas dataframe

    Outputs:
        - None, all absences from the list are saved to Outlook calendar
    """
    # initialize Outlook application
    objOutlook = win32.Dispatch('Outlook.Application')

    for intAbsence in range(len(plstAbsenceData)):
        # convert dates to datetime
        dttStart = ggf.dttConvertDate(plstAbsenceData[intAbsence][0], '-')
        dttEnd = ggf.dttConvertDate(plstAbsenceData[intAbsence][1], '-')

        # set next absence to compare with the current
        if intAbsence == (len(plstAbsenceData) - 1):
            lstAbsence2 = None
        else:
            lstAbsence2 = plstAbsenceData[intAbsence + 1]

        # determine correct half day
        intHalfDay = intDetermineHalfDay(
            plstAbsenceData[intAbsence],
            lstAbsence2
        )

        # calculate absence duration, convert from seconds to days
        if plstAbsenceData[intAbsence][2] == 0.5:
            # for half days, use explicit half day duration
            fltDuration = 0.5
        else:
            # for multiday absences calculate the absence duration from dates
            dttDuration = dttEnd - dttStart + datetime.timedelta(days = 1)
            fltDuration = dttDuration.total_seconds() / 60 / 60 / 24

        # save the absence to Outlook
        SaveAbsence(
            objOutlook,
            dttStart,
            fltDuration,
            plstAbsenceData[intAbsence][4],
            g.INT_MEETING_OUT_OF_OFFICE,
            intHalfDay
        )

def SaveAbsence(
    pobjApplication,
    pdttFrom,
    pfltDurationDays,
    pstrType,
    pintStatus,
    pintHalfDay = 1
):
    """Save absence to Outlook calendar.

    Inputs:
        - pobjApplication - instance of Outlook application
        - pdttFrom - datetime format of absence start date
        - pfltDurationDays - length of absence duration in days, float
        - pstrType - string, literal name of the absence type, eg 'Vacation'
        - pintStatus - numeric representation of appointment status in Outlook
        convention
        - pintHalfDay - optional, indicates in which half of the day should
        the absence start, 1 = first half of the day until 12:00
    """
    # create new appointment
    objAbsence = pobjApplication.CreateItem(1)

    if int(pfltDurationDays) != pfltDurationDays:
        # half day absence, set it up correctly
        if pintHalfDay == 1:
            # starts in the morning
            strTime = '7:00'
        else:
            # starts in the afternoon
            strTime = '12:00'

    else:
        # full day absence, set a default start
        strTime = '0:00'

    # set start date and time
    objAbsence.Start = pdttFrom.strftime('%Y-%m-%d') + ' ' + strTime

    # set end time using duration in minutes
    # for calendar view purpose, change first half day absence duration
    if pintHalfDay == 1 and pfltDurationDays < 1:
        # make half-day absence in the first half last for 5 hours
        objAbsence.Duration = 60 * 5
    else:
        # for all other absences, the view is OK with the following calculation
        objAbsence.Duration = 60 * 24 * pfltDurationDays

    # add subject
    objAbsence.Subject = pstrType + g.STR_ABSENCE_CALENDAR_NAME

    # set absence status
    objAbsence.BusyStatus = pintStatus

    # save the absence
    objAbsence.Save()

def ImportAbsences():
    """Coordinate import of source data and submission to Outlook calendar.

    Inputs:
        - None

    Outputs:
        - None, absences from the source file are saved as appointments in the
        default Outlook calendar
    """
    if os.path.isfile(g.STR_FULL_PATH_SCRAPED_DATA):
        # read the data to import
        dtfSource = pd.read_csv(g.STR_FULL_PATH_SCRAPED_DATA, sep = '\t')

        # check if the data is relevant for importing
        if blnVerifyData(dtfSource):
            # data is relevant, import it for processing
            lstProcessed = lstSplitOffHalfDays(dtfSource)

            # import the processed data to Outlook calendar
            SaveAbsences(lstProcessed)
        else:
            # data not relevant, inform the user
            strMessage = 'Nothing relevant to import in the source data'
            print(strMessage)

    else:
        # if import failed, inform the user
        strMessage = 'Import of absences to Outlook failed, unable to import '
        strMessage += os.path.abspath(g.STR_FULL_PATH_SCRAPED_DATA)
        print(strMessage)