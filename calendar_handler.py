# %%
# Contains methods and functions to acquire and output data from Outlook 
# calendar

# %% import modules
import win32com.client
import getpass
import datetime
import pywintypes

import globals as g

# %% define functions
def objCreateRecipient():
    '''
    Create Outlook recipient object using the email address of the user logged
    in Windows

    Inputs:
        - None

    Outputs:
        - objRecipient - Outlook recipient object
    '''
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
    '''
    Get calendar status in Outlook API form for the specified recipient
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
    '''
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
    '''
    Converts YYYYMMDD string date to datetime object

    Inputs:
        - pstrYYYYMMDD - string date in YYYYMMDD format

    Outputs:
        - dttConverted - converted date as a datetime object
    '''
    # parse the data from string and convert to datetime
    dttConverted = datetime.datetime(
        int(pstrYYYYMMDD[:4]),
        int(pstrYYYYMMDD[4:6]),
        int(pstrYYYYMMDD[6:])
    )

    return dttConverted

def strConvertDate(pdttDateTime):
    '''
    Converts datetime object into date in YYYYMMDD format

    Inputs:
        - pdttDateTime - datetime object

    Outputs:
        - strConverted - string representation of a date in YYYYMMDD format
    '''
    # convert date to YYYYMMDD date
    strConverted = pdttDateTime.strftime('%Y%m%d')

    return strConverted

def strConvertAbsence(pintAbsenceCode, pblnXperience = False):
    '''
    Converts Outlook absence code to word representation corresponding with the
    selected output type

    Inputs:
        - pintAbsenceCode - Outlook API numeric code of meeting status
        - pblnXperience - optional boolean indicator if the values should be
        converted to Outlook or Xperience naming, default output is Outlook

    Outputs:
        - strAbsence - word representation of the corresponding meeting status
    '''
    if pblnXperience:
        # use xperience naming
        if pintAbsenceCode == g.INT_MEETING_FREE:
            # working from home
            strAbsence = g.STR_ABSENCE_TYPE_HOME_OFFICE
        elif pintAbsenceCode == g.INT_MEETING_TENTATIVE:
            strAbsence = g.STR_MEETING_TENTATIVE
        elif pintAbsenceCode == g.INT_MEETING_BUSY:
            strAbsence = g.STR_MEETING_BUSY
        elif pintAbsenceCode == g.INT_MEETING_OUT_OF_OFFICE:
            strAbsence = g.STR_MEETING_OUT_OF_OFFICE
        elif pintAbsenceCode == g.INT_MEETING_WORKING_ELSEWHERE:
            # working from office
            strAbsence = g.STR_ABSENCE_TYPE_NONE
        else:
            strAbsence = 'Unknown absence code'
    else:
        # use standard outlook naming
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
    '''
    Retrieves all calendar statuses in a given time period

    Inputs:
        - pstrDateStart - start date of the analyzed period in YYYYMMDD format,
        included
        - pstrDateEnd - end date of the analyzed period in YYYYMMDD format,
        included

    Outputs:
        - lstStatuses - list of 24 hour calendar meeting statuses from the
        requested period
    '''
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
    '''
    Analyzes a stream of calendar statuses on full-day basis. Returns overall
    daily status based on all-day meeting types. If no full day meeting is
    present, free status is returned.

    Inputs:
        - pstrStatus - single daily status coming from the function that returns
        24 hours of meeting availability from calendar

    Outputs:
        - intFullDayStatus - status of an entire day, if no full day meeting is
        in the calendar, free status is returned
    '''
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
    '''
    Returns a list of tuples containing daily calendar status for every day in
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
    '''
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

def lstAggregateCalendarOutput(plstDailyStatus):
    '''
    Aggregates consecutive days of same status into a single record

    Inputs:
        - plstDailyStatus - list of tuples containing start date, end date and
        calendar status for the given period. Dates are stored as datetime
        objects, status is an Outlook API constant

    Outputs:
        - lstStatuses - list of aggregated dates
    '''
    # to avoid modifying the original list, clone the input
    lstStatuses = plstDailyStatus.copy()

    # initialize a looping variable
    intAggregate = 0

    # set the stopping bound (moving)
    intUntil = len(lstStatuses)

    # aggregate the data
    while intAggregate + 1 < intUntil:
        if lstStatuses[
            intAggregate
        ][1] + datetime.timedelta(days = 1) == lstStatuses[
            intAggregate + 1
        ][0] \
        and lstStatuses[
            intAggregate
        ][2] == lstStatuses[intAggregate + 1][2]:
            # two neighbouring dates with mathcing status
            # merge into one record as a tuple
            lstStatuses[intAggregate] = lstStatuses[intAggregate][0], \
                lstStatuses[intAggregate+1][0], \
                lstStatuses[intAggregate][2]

            # remove the merged observation
            lstStatuses.pop(intAggregate + 1)

            # shorten the loop length
            intUntil = len(lstStatuses)
        else:
            # current pair of records can't be merged, move to the next one
            intAggregate += 1

    return lstStatuses

def lstConvertAggregatedOutput(plstAggregatedData, pblnXperienceOutput = False):
    '''
    Converts list of aggregated data with time period and corresponding statuses
    to a human-readable form containing dates in DD/MM/YYYY format and word
    descriptions of the availability statuses

    Inputs:
        - plstAggregatedData - list of tuples containing start date, end date,
        and status. Dates are datetime objects, status is an Outlook constant
        - pblnXperienceOutput - optional boolean indicator to specify if the
        output should be converted to output submittable to xperience, default
        value is set to false (= Outlook output chosen instead)

    Outputs:
        - lstStringOutput - list of tuples containing start, end dates in
        DD/MM/YYYY formats, and calendar all day status either in Outlook or
        Xperience format
    '''
    # initialize a list for outputs
    lstStringOutput = []

    # process dates and statuses to human readable form
    for tplEntry in plstAggregatedData:
        # convert the start date and add a tab
        strConvertedEntry = tplEntry[0].strftime('%d/%m/%Y')
        strConvertedEntry += '\t'

        # convert the end date and add a tab
        strConvertedEntry += tplEntry[1].strftime('%d/%m/%Y')
        strConvertedEntry += '\t'

        # convert status and add a line break
        strConvertedEntry += strConvertAbsence(tplEntry[2], pblnXperienceOutput)
        strConvertedEntry += '\n'

        # append the converted entry to the list
        lstStringOutput.append(strConvertedEntry)

    return lstStringOutput

def OutputCalendarData(plstStringData):
    '''
    Stores the provided data in a text file.

    Inputs:
        - plstStringData - list of data rows to be saved

    Outputs:
        - None, a new text file is created to contain the provided data
    '''
    # open a new text file for writing
    objOutput = open(g.STR_FULL_PATH_CALENDAR_DATA, 'w')

    # write contents of the list into the file
    objOutput.writelines(plstStringData)

    # close the file
    objOutput.close()

# %% run the process
def AnalyzeCalendar(pstrDateStart, pstrDateEnd, pblnXperience = False):
    '''
    Runs the process of the calendar analysis in the requested time period.
    Reads the calendar data from own Outlook calendar, aggregates the output
    based on the absence types and converts the results to a readable output.
    The output is then saved to a text file.

    Inputs:
        - pstrDateStart - start date of the calendar analysis in YYYYMMDD format
        - pstrDateEnd - end date of the calendar analysis in YYYYMMDD format
        - pblnXperience - boolean indicator if the output should be in Xperience
        or standard format, default set to standard

    Outputs:
        - None, a text file with the calendar data in the requested format is
        created
    '''
    # retrieve all 
    lstCalendar = lstGetFullDayOutputInPeriod(
        pstrDateStart,
        pstrDateEnd
    )

    # aggregate the days with the same output
    lstCalendar = lstAggregateCalendarOutput(lstCalendar)

    # convert the calendar info to Xperience output
    lstCalendar = lstConvertAggregatedOutput(lstCalendar, pblnXperience)

    # output the calendar data to a file
    OutputCalendarData(lstCalendar)