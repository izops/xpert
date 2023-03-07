# %% import modules
import re
import win32com.client
import getpass
import datetime
import pywintypes

import globals as g

# %% define functions
def objCreateRecipient():
    # initialize Outlook application and connect to its namespace
    objOutlook = win32com.client.Dispatch('Outlook.Application')
    objNamespace = objOutlook.GetNamespace('MAPI')

    # get own email address based on the user name
    strEmail = getpass.getuser() + g.STR_USER_DOMAIN

    # create a recipient object based on own email address
    objRecipient = objNamespace.CreateRecipient(strEmail)

    return objRecipient

def strGetStatus(
    pobjRecipient,
    pstrYYYYMMDD,
    pintTimeInterval
):
    # put together a datetime value for checking the status
    # using arbitrary time to get data from the correct date
    dttStartDate = dttConvertDate(pstrYYYYMMDD) + datetime.timedelta(hours = 1)

    # convert source time to pywin time
    objPywinTime = pywintypes.Time(dttStartDate)

    # retrieve status of the calendar
    strCalendarStatus = pobjRecipient.FreeBusy(
        objPywinTime,
        pintTimeInterval,
        True
    )

    return strCalendarStatus

def dttConvertDate(pstrYYYYMMDD):
    dttConverted = datetime.datetime(
        int(pstrYYYYMMDD[:4]),
        int(pstrYYYYMMDD[4:6]),
        int(pstrYYYYMMDD[6:])
    )

    return dttConverted

def strConvertDate(pdttDateTime):
    # convert date to YYYYMMDD date
    strConverted = pdttDateTime.strftime('%Y%m%d')

    return strConverted

def strConvertAbsence(pintAbsenceCode):
    if pintAbsenceCode == g.INT_MEETING_FREE:
        strAbsence = g.STR_MEETING_FREE
    elif pintAbsenceCode == g.INT_MEETING_TENTATIVE:
        strAbsence = g.STR_MEETING_TENTATIVE
    elif pintAbsenceCode == g.INT_MEETING_BUSY:
        strAbsence = g.STR_MEETING_BUSY
    elif pintAbsenceCode == g.INT_MEETING_OUT_OF_OFFICE:
        strAbsence = g.STR_MEETING_OUT_OF_OFFICE
    elif pintAbsenceCode == g.INT_MEETING_WORKING_ELSEWHERE:
        strAbsence = g.STR_MEETING_WORKING_ELSEWHERE
    else:
        strAbsence = 'Unknown absence code'

    return strAbsence

def lstGetCalendarStatuses(pstrDateStart, pstrDateEnd):
    # create Outlook recipient
    objRecipient = objCreateRecipient()

    # convert the date to the datetime value
    dttDateCurrent = dttConvertDate(pstrDateStart)
    dttDateEnd = dttConvertDate(pstrDateEnd)

    # initialize a list of calendar statuses
    lstStatuses = []

    # get calendar statuses for each day in the start-end date range
    while dttDateCurrent <= dttDateEnd:
        # get status of the calendar for the day by hours
        strStatus = strGetStatus(
            objRecipient,
            strConvertDate(dttDateCurrent),
            g.INT_CALENDAR_TIME_UNIT_HOURS
        )

        # keep only first 24 hours from the output
        strStatus = strStatus[:24]

        # save the statuses
        lstStatuses.append(strStatus)

        # go to the next day
        dttDateCurrent += datetime.timedelta(days = 1)

    return lstStatuses

def intAnalyzeCalendarStatus(pstrStatus):
    # check the all day statuses against the Outlook values
    if pstrStatus == str(g.INT_MEETING_OUT_OF_OFFICE) * 24:
        # full day out of office
        intFullDayStatus = g.INT_MEETING_OUT_OF_OFFICE
    elif pstrStatus == str(g.INT_MEETING_WORKING_ELSEWHERE) * 24:
        # full day working elsewhere
        intFullDayStatus = g.INT_MEETING_WORKING_ELSEWHERE
    elif pstrStatus == str(g.INT_MEETING_BUSY) * 24:
        # full day busy
        intFullDayStatus = g.INT_MEETING_BUSY
    else:
        # not a full day meeting, set to free
        intFullDayStatus = g.INT_MEETING_FREE

    return intFullDayStatus

def lstGetFullDayOutputInPeriod(
    pstrPeriodStart,
    pstrPeriodEnd,
    blnRemoveWeekend = True
):
    # get calendar statuses for every day
    lstCalendarStatuses = lstGetCalendarStatuses(pstrPeriodStart, pstrPeriodEnd)

    # convert the start date to datetime
    dttCurrentDay = dttConvertDate(pstrPeriodStart)

    # initialize a list of full day output in the period
    lstCalendarPeriod = []

    # get status for each day of the period
    for strCurrentStatus in lstCalendarStatuses:
        if (dttCurrentDay.weekday() < 5 and blnRemoveWeekend) \
        or not blnRemoveWeekend:
            # get the all day status of the day
            intAllDayStatus = intAnalyzeCalendarStatus(strCurrentStatus)

            # # fix the case when there is no all day meeting
            # if intAllDayStatus < 0:
            #     # special case of no all day meeting
            #     intAllDayStatus = g.INT_MEETING_FREE
            
            # store the date and status
            tplDailyStatus = (dttCurrentDay, intAllDayStatus)

            # store it in the list
            lstCalendarPeriod.append(tplDailyStatus)

        # increment day to the next one in the period
        dttCurrentDay += datetime.timedelta(days = 1)
    
    return lstCalendarPeriod
