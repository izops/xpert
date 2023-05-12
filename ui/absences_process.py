# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import web.credentials as wcr
import web.submit_absences as wsa
import web.scrape_absences as wsc
import web.process_scraped_data as wps
import ui.absences_inputs as uai
import ui.scraping_inputs as usi
import calendar_works.calendar_analyzer as cca

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
        strDateStart = uai.strGetDate(g.STR_UI_REQUEST_DATE_START)
        strDateEnd = uai.strGetDate(g.STR_UI_REQUEST_DATE_END)

    if pintChoice in lstAbsenceSubmission:
        # get the calendar convention
        strConvention = uai.strGetCalendarConvention()

        # create boolean parameter based on the output from convention check
        if strConvention == g.LST_UI_ANSWERS_CONVENTION[0]:
            blnOfficeFocused = True
        else:
            blnOfficeFocused = False

    if pintChoice in lstScraping:
        # get type of absence to scrape
        strScrapeAbsence = usi.strGetAbsenceType

    if pintChoice not in lstNoCredentials and strScrapeAbsence != 'c':
        # get xperience credentials
        strUserName = wcr.strGetUserName()
        strPassword = wcr.strGetPassword()

    if pintChoice in lstAbsenceSubmission:
        # launch the calendar analysis
        cca.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

        # inform the user about the process end
        print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

    if pintChoice in [
        g.INT_UI_CHOICE_XPERIENCE_SUBMISSION,
        g.INT_UI_CHOICE_FULL_SUBMISSION
    ]:
        # run absence submission process
        wsa.SubmitAbsences(strUserName, strPassword)

        # discard the password
        del strPassword

    if pintChoice in [
        g.INT_UI_CHOICE_XPERIENCE_SCRAPE,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ] and strScrapeAbsence != 'c':
        # scrape xperience data, process and save them to an external file
        wps.ObtainXperienceAbsences(
            strUserName,
            strPassword,
            strDateStart,
            strDateEnd,
            strScrapeAbsence
        )

        # discard the password
        del strPassword

    if pintChoice in [
        g.INT_UI_CHOICE_OUTLOOK_IMPORT,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ] and strScrapeAbsence != 'c':
        # save xperience data to outlook - TO BE DONE
        pass

    # the process ends here, say goodbye to the user
    print(g.STR_UI_GOODBYE)