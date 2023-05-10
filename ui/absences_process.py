# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import web.credentials as c
import web.submit_absences as w
import ui.absences_inputs as abi
import calendar_works.calendar_analyzer as ca

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
    if pintChoice != g.INT_UI_CHOICE_XPERIENCE_SUBMISSION:
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

    if pintChoice != g.INT_UI_CHOICE_OUTLOOK_EXPORT:
        # get xperience credentials
        strUserName = c.strGetUserName()
        strPassword = c.strGetPassword()

    if pintChoice != g.INT_UI_CHOICE_XPERIENCE_SUBMISSION:
        # launch the calendar analysis
        ca.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

    if pintChoice != g.INT_UI_CHOICE_OUTLOOK_EXPORT:
        # run xperience process
        w.SubmitAbsences(strUserName, strPassword)

        # discard the password
        del strPassword

    # the process ends here, say goodbye to the user
    print(g.STR_UI_GOODBYE)