# %% import modules
import sys
import globals as g
import calendar_handler as c

sys.path.append('../emea_oth_xpert')
import web.submit_absences as w
import ui.absences_inputs as abi

# %% define the master method to launch the process parts
def RunProcess(pintChoice):
    """Based on the input, run Outlook calendar analysis, submission of
    absences to Xperience or both.

    Inputs:
        - pintChoice - numeric indication of the process to be run

    Outputs:
        - None, either one or two processes are run
    """
    # analyze the calendar
    if pintChoice in [g.INT_UI_CHOICE_CALENDAR, g.INT_UI_CHOICE_ALL]:
        # request the starting and ending point for the calendar analysis
        strDateStart = abi.strGetDate(g.STR_UI_REQUEST_DATE_START)
        strDateEnd = abi.strGetDate(g.STR_UI_REQUEST_DATE_END)

        # get the calendar convention
        strConvention = abi.strGetCalendarConvention()

        # create boolean parameter based on the output from convention check
        if strConvention == g.LST_UI_ANSWERS_CONVENTION[0]:
            blnOfficeFocused = True
        else:
            blnOfficeFocused = False

        # launch the calendar analysis
        c.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

        if pintChoice == g.INT_UI_CHOICE_CALENDAR:
            # the process ends here, say goodbye to the user
            print(g.STR_UI_GOODBYE)
        
    if pintChoice in [g.INT_UI_CHOICE_XPERIENCE, g.INT_UI_CHOICE_ALL]:
        # run xperience process
        w.SubmitAbsences()

        # the process ends here, say goodbye to the user
        print(g.STR_UI_GOODBYE)