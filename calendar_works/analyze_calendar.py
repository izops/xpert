# %% import modules
import datetime
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define functions and methods to analyze calendar data
def strConvertAbsence(pintAbsenceCode, pblnOfficeFocused = True):
    """Convert Outlook absence code to word representation corresponding
    with the selected output type.

    Inputs:
        - pintAbsenceCode - Outlook API numeric code of meeting status
        - pblnOfficeFocused - optional boolean indicator if the values should
        be converted with 'working elsewhere' used as home office (True) or as
        working from home (False)

    Outputs:
        - strAbsence - word representation of the corresponding meeting status
    """
    # log inputs
    logging.info('strConvertAbsence - pintAbsenceCode: ' + str(
        pintAbsenceCode
    ))
    logging.info('strConvertAbsence - pblnOfficeFocused: ' + str(
        pblnOfficeFocused
    ))

    if pblnOfficeFocused:
        # use office-focused naming
        if pintAbsenceCode == g.INT_MEETING_FREE:
            # working from office
            strAbsence = g.STR_ABSENCE_TYPE_NONE
        elif pintAbsenceCode == g.INT_MEETING_TENTATIVE:
            strAbsence = g.STR_MEETING_TENTATIVE
        elif pintAbsenceCode == g.INT_MEETING_BUSY:
            strAbsence = g.STR_MEETING_BUSY
        elif pintAbsenceCode == g.INT_MEETING_OUT_OF_OFFICE:
            strAbsence = g.STR_MEETING_OUT_OF_OFFICE
        elif pintAbsenceCode == g.INT_MEETING_WORKING_ELSEWHERE:
            # working from home
            strAbsence = g.STR_ABSENCE_TYPE_HOME_OFFICE
        elif pintAbsenceCode == g.INT_MEETING_MIXED:
            # mixed absences
            strAbsence = g.STR_MEETING_MIXED
        else:
            strAbsence = 'Unknown absence code'
    else:
        # use home-office-focused naming
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

    return strAbsence


def intAnalyzeCalendarStatus(pstrStatus):
    """Analyze a stream of calendar statuses on full-day basis. Return overall
    daily status based on all-day meeting types. If no full day meeting is
    present, free status is returned.

    Inputs:
        - pstrStatus - single daily status coming from the function that
        returns 24 hours of meeting availability from calendar

    Outputs:
        - intFullDayStatus - status of an entire day, if no full day meeting is
        in the calendar, free status is returned
    """
    # log inputs
    logging.info('intAnalyzeCalendarStatus - pstrStatus: ' + pstrStatus)
    
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
        # not a full day meeting, verify if it contains partial out of office
        # or working from elsewhere
        if str(g.INT_MEETING_OUT_OF_OFFICE) in pstrStatus \
        or str(g.INT_MEETING_WORKING_ELSEWHERE) in pstrStatus:
            # there is partial absence in the day, set it to mixed
            intFullDayStatus = g.INT_MEETING_MIXED
        else:
            # no issues with the day, set it to free all day status
            intFullDayStatus = g.INT_MEETING_FREE

    return intFullDayStatus

def lstAggregateCalendarOutput(plstDailyStatus):
    """Aggregate consecutive days of same status into a single record.

    Inputs:
        - plstDailyStatus - list of tuples containing start date, end date and
        calendar status for the given period. Dates are stored as datetime
        objects, status is an Outlook API constant

    Outputs:
        - lstStatuses - list of aggregated dates
    """
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

def lstConvertAggregatedOutput(plstAggregatedData, pblnOfficeFocused = True):
    """Convert list of aggregated data with time period and corresponding
    statuses to a human-readable form containing dates in DD/MM/YYYY format 
    and word descriptions of the availability statuses.

    Inputs:
        - plstAggregatedData - list of tuples containing start date, end date,
        and status. Dates are datetime objects, status is an Outlook constant
        - pblnOfficeFocused - optional boolean indicator to specify calendar
        convention for working elsewhere as either home office (True) or work
        from office (False)

    Outputs:
        - lstStringOutput - list of tuples containing start, end dates in
        DD/MM/YYYY formats, and calendar all day status either in Outlook or
        Xperience format
    """
    # initialize a list for outputs
    lstStringOutput = []

    # process dates and statuses to human readable form
    for tplEntry in plstAggregatedData:
        # convert the start date and add a tab
        strConvertedEntry = tplEntry[0].strftime(g.STR_DATE_FORMAT)
        strConvertedEntry += '\t'

        # convert the end date and add a tab
        strConvertedEntry += tplEntry[1].strftime(g.STR_DATE_FORMAT)
        strConvertedEntry += '\t'

        # convert status and add a line break
        strConvertedEntry += strConvertAbsence(tplEntry[2], pblnOfficeFocused)
        strConvertedEntry += '\n'

        # append the converted entry to the list
        lstStringOutput.append(strConvertedEntry)

    return lstStringOutput

def OutputCalendarData(plstStringData):
    """Store the provided data in a text file.

    Inputs:
        - plstStringData - list of data rows to be saved

    Outputs:
        - None, a new text file is created to contain the provided data
    """
    # open a new text file for writing
    objOutput = open(g.STR_FULL_PATH_CALENDAR_DATA, 'w')

    # write contents of the list into the file
    objOutput.writelines(plstStringData)

    # close the file
    objOutput.close()