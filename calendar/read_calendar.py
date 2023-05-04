# %% import modules

# %% define functions and methods required for analysis of existing absences
def objCreateRecipient():
    """Create Outlook recipient object using the email address of the user
    logged in Windows.

    Inputs:
        - None

    Outputs:
        - objRecipient - Outlook recipient object
    """
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
    """Get calendar status in Outlook API form for the specified recipient
    starting by a specific date at midnight in the local timezone. Each
    character of the return string represents a unit of time interval in minutes
    defined in input.

    Inputs:
        - pobjRecipient - Outlook recipient object to read calendar data from
        - pstrYYYYMMDD - start date from which the statuses will be returned
        - pintTimeInterval - time interval in minutes by which the calendar will
        be checked. Eg to check calendar status by hour, use time interval = 60

    Outputs:
        - strCalendarStatus - characters representing time interval calendar
        availability from requested time up to 30 days
    """
    # put together a datetime value for checking the status
    # using arbitrary time to get data from the correct date
    dttStartDate = dttConvertDate(pstrYYYYMMDD) + datetime.timedelta(hours = 12)

    # convert source time to pywin time
    objPywinTime = pywintypes.Time(dttStartDate)

    # retrieve status of the calendar
    strCalendarStatus = pobjRecipient.FreeBusy(
        objPywinTime,
        pintTimeInterval,
        True
    )

    return strCalendarStatus

def lstGetCalendarStatuses(pstrDateStart, pstrDateEnd):
    """Retrieve all calendar statuses in a given time period.

    Inputs:
        - pstrDateStart - start date of the analyzed period in YYYYMMDD format,
        included
        - pstrDateEnd - end date of the analyzed period in YYYYMMDD format,
        included

    Outputs:
        - lstStatuses - list of 24 hour calendar meeting statuses from the
        requested period
    """
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

def lstGetFullDayOutputInPeriod(
    pstrPeriodStart,
    pstrPeriodEnd,
    blnRemoveWeekend = True
):
    """Return a list of tuples containing daily calendar status for every day in
    the specified period, including the start and the end days. It is possible
    to exclude weekend days from the status analysis.

    Inputs:
        - pstrPeriodStart - start date of the analyzed period, in YYYYMMDD
        format, included in the period
        - pstrPeriodEnd - end  date of the analyzed period, in YYYYMMDD format,
        included in the period

    Outputs:
        - lstCalendarPeriod - list of tuples containing start date, end date,
        and calendar status. Dates are returned as datetime objects and the
        status is an Outlook API constant
    """
    # get calendar statuses for every day
    lstCalendarStatuses = lstGetCalendarStatuses(pstrPeriodStart, pstrPeriodEnd)

    # convert the start date to datetime
    dttCurrentDay = dttConvertDate(pstrPeriodStart)

    # initialize a list of full day output in the period
    lstCalendarPeriod = []

    # get status for each day of the period
    for strCurrentStatus in lstCalendarStatuses:
        # continue only if the day is either weekday or all days are requested
        if (dttCurrentDay.weekday() < 5 and blnRemoveWeekend) \
        or not blnRemoveWeekend:
            # get the all day status of the day
            intAllDayStatus = intAnalyzeCalendarStatus(strCurrentStatus)
            
            # store the date from, date to and status
            tplDailyStatus = (dttCurrentDay, dttCurrentDay, intAllDayStatus)

            # store it in the list
            lstCalendarPeriod.append(tplDailyStatus)

        # increment day to the next one in the period
        dttCurrentDay += datetime.timedelta(days = 1)
    
    return lstCalendarPeriod