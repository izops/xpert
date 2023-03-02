# %% import modules
import re
import win32com.client
import getpass
import datetime
import pywintypes

import globals as g

# %% define functions
# open Outlook session

def objCreateRecipient():
    # initialize Outlook application and connect to its namespace
    objOutlook = win32com.client.Dispatch('Outlook.Application')
    objNamespace = objOutlook.GetNamespace('MAPI')

    # get own email address based on the user name
    strEmail = getpass.getuser + g.STR_USER_DOMAIN

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

def lstGetStatuses(pstrDateStart, pstrDateEnd):
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
            objCreateRecipient,
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