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


# def SaveAbsences(pdtfAbsenceData):
#     # convert the absences to a list
#     lstAbsences = pdtfAbsenceData.values.tolist()

#     for intAbsence in len(lstAbsences):
#         # attempt to append half day absence to the full day
#         if lstAbsences[intAbsence][2] == 0.5:
#             # check if the half day starts a chunk of absences
#             if intAbsence < (len(lstAbsences) - 1):
#                 if lstAbsences[intAbsence + 1][0] == datetime.timedelta()


def SaveAbsence(
    pobjApplication,
    pdttFrom,
    pdblDurationDays,
    pstrType,
    pintStatus,
    pintHalfDay = 1
):
    """Save absence to Outlook calendar.

    Inputs:
        - pobjApplication - instance of Outlook application
        - pdttFrom - datetime format of absence start date
        - pdblDurationDays - length of absence duration in days, double
        - pstrType - string, literal name of the absence type, eg 'Vacation'
        - pintStatus - numeric representation of appointment status in Outlook
        convention
        - pintHalfDay - optional, indicates in which half of the day should
        the absence start, 1 = first half of the day until 12:00
    """
    # create new appointment
    objAbsence = pobjApplication.CreateItem(1)

    if int(pdblDurationDays) != pdblDurationDays:
        # half day absence, set it up correctly
        if pintHalfDay == 1:
            # starts in the morning
            strTime = '0:00'
        else:
            # starts in the afternoon
            strTime = '12:00'

    else:
        # full day absence, set a default start
        strTime = '0:00'

    # set start date and time
    objAbsence.Start = pdttFrom.strftime('%Y-%m-%d') + ' ' + strTime

    # set end time using duration in minutes
    objAbsence.Duration = 60 * 24 * pdblDurationDays

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