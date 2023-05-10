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
    # group choices based on flowchart
    lstAbsenceSubmission = [
        g.INT_UI_CHOICE_OUTLOOK_EXPORT,
        g.INT_UI_CHOICE_FULL_SUBMISSION
    ]

    lstScraping = [
        g.INT_UI_CHOICE_XPERIENCE_SCRAPE,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ]

    lstNoCredentials = [
        g.INT_UI_CHOICE_OUTLOOK_EXPORT,
        g.INT_UI_CHOICE_OUTLOOK_IMPORT
    ]

    # coordinate the process steps based on user choice
    if pintChoice in lstAbsenceSubmission or pintChoice in lstScraping:
        # request the starting and ending point for the calendar analysis
        strDateStart = abi.strGetDate(g.STR_UI_REQUEST_DATE_START)
        strDateEnd = abi.strGetDate(g.STR_UI_REQUEST_DATE_END)

    if pintChoice in lstAbsenceSubmission:
        # get the calendar convention
        strConvention = abi.strGetCalendarConvention()

        # create boolean parameter based on the output from convention check
        if strConvention == g.LST_UI_ANSWERS_CONVENTION[0]:
            blnOfficeFocused = True
        else:
            blnOfficeFocused = False

    if pintChoice in lstScraping:
        # get type of absence to scrape - TO BE DONE
        pass

    if pintChoice not in lstNoCredentials:
        # get xperience credentials
        strUserName = c.strGetUserName()
        strPassword = c.strGetPassword()

    if pintChoice in lstAbsenceSubmission:
        # launch the calendar analysis
        ca.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

    if pintChoice in [
        g.INT_UI_CHOICE_XPERIENCE_SUBMISSION,
        g.INT_UI_CHOICE_FULL_SUBMISSION
    ]:
        # run absence submission process
        w.SubmitAbsences(strUserName, strPassword)

        # discard the password
        del strPassword

    if pintChoice in [
        g.INT_UI_CHOICE_XPERIENCE_SCRAPE,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ]:
        # scrape xperience data - TO BE DONE
        pass

    if pintChoice in [
        g.INT_UI_CHOICE_OUTLOOK_IMPORT,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ]:
        # save xperience data to outlook - TO BE DONE
        pass

    # the process ends here, say goodbye to the user
    print(g.STR_UI_GOODBYE)