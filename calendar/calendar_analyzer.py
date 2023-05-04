# %%
# Contains methods and functions to acquire and output data from Outlook 
# calendar

# %% import modules
import globals as g

# %% define calendar analysis method
def AnalyzeCalendar(pstrDateStart, pstrDateEnd, pblnOfficeFocused = True):
    """Run the process of the calendar analysis in the requested time period.
    Read the calendar data from own Outlook calendar, aggregate the output
    based on the absence types and convert the results to a readable output.
    The output is then saved to a text file.

    Inputs:
        - pstrDateStart - start date of the calendar analysis in YYYYMMDD format
        - pstrDateEnd - end date of the calendar analysis in YYYYMMDD format
        - pblnOfficeFocused - optional boolean indicator to specify the calendar
        convention for working elsewhere as either home office (True) or work
        from office (False)

    Outputs:
        - None, a text file with the calendar data in the requested format is
        created
    """
    # retrieve all 
    lstCalendar = lstGetFullDayOutputInPeriod(
        pstrDateStart,
        pstrDateEnd
    )

    # aggregate the days with the same output
    lstCalendar = lstAggregateCalendarOutput(lstCalendar)

    # convert the calendar info to Xperience output
    lstCalendar = lstConvertAggregatedOutput(lstCalendar, pblnOfficeFocused)

    # output the calendar data to a file
    OutputCalendarData(lstCalendar)