# %% import modules
import logging
import sys

sys.path.append('../emea_oth_xpert')
import general.global_constants as g
import web.credentials as wcr
import web.submit_absences as wsa
import web.process_scraped_data as wps
import ui.absences_inputs as uai
import ui.scraping_inputs as usi
import calendar_works.calendar_analyzer as cca
import calendar_works.absence_import as cai

# %% set up logging
logging.basicConfig(
    level = g.OBJ_LOGGING_LEVEL,
    format=' %(asctime)s -  %(levelname)s -  %(message)s'
)

# %% define the master method to launch the process parts
def RunProcess(pintChoice):
    """Based on the input, run Outlook calendar analysis, submission of
    absences to Xperience or both.

    Inputs:
        - pintChoice - numeric indication of the process to be run

    Outputs:
        - None, either one or two processes are run
    """    
    # log input
    logging.info('RunProcess - pintChoice: ' + str(pintChoice))

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

        # log obtained data
        logging.debug('RunProcess - strDateStart: ' + strDateStart)
        logging.debug('RunProcess - strDateEnd: ' + strDateEnd)

    if pintChoice in lstAbsenceSubmission:
        # get the calendar convention
        strConvention = uai.strGetCalendarConvention()

        # log obtained data
        logging.debug('RunProcess - strConvention: ' + strConvention)

        # create boolean parameter based on the output from convention check
        if strConvention == g.LST_UI_ANSWERS_CONVENTION[0]:
            blnOfficeFocused = True
        else:
            blnOfficeFocused = False

        # log convention flag
        logging.debug('RunProcess - blnOfficeFocused: ' + str(
            blnOfficeFocused
        ))

    if pintChoice in lstScraping:
        # get type of absence to scrape
        strScrapeAbsence = usi.strGetAbsenceType()
    else:
        # set scrape value to default if not selected
        strScrapeAbsence = ''

    # log absence type to scrape
    logging.debug('RunProcess - strScrapeAbsence: ' + strScrapeAbsence)

    if pintChoice not in lstNoCredentials \
    and strScrapeAbsence != 'c' \
    and pintChoice >= 0:
        # get xperience credentials
        strUserName = wcr.strGetUserName()
        strPassword = wcr.strGetPassword()

    # initialize safety parameter
    blnSafety = True

    if pintChoice in lstAbsenceSubmission:
        # inform user about the process start
        print(g.STR_UI_PROCESS_CALENDAR_ANALYSIS + g.STR_UI_PROCESS_STARTED)
        
        try:
            # launch the calendar analysis
            cca.AnalyzeCalendar(strDateStart, strDateEnd, blnOfficeFocused)

            # inform the user about the process end
            print(g.STR_UI_CALENDAR_ANALYSIS_COMPLETE)

        except:
            # process failed, inform the user
            print(g.STR_UI_PROCESS_CALENDAR_ANALYSIS + g.STR_UI_PROCESS_FAILED)

            # change the safety indicator
            blnSafety = False

    if blnSafety and pintChoice in [
        g.INT_UI_CHOICE_XPERIENCE_SUBMISSION,
        g.INT_UI_CHOICE_FULL_SUBMISSION
    ]:
        # inform user about the process start
        print(g.STR_UI_PROCESS_XPERIENCE_SUBMISSION + g.STR_UI_PROCESS_STARTED)

        try:
            # run absence submission process
            blnContinue = wsa.blnSubmitAbsences(strUserName, strPassword)

            # inform the user about the process end if ended successfully
            if blnContinue:
                print(g.STR_UI_SUBMISSION_TO_XPERIENCE)

        except:
            # process failed, inform the user
            strMessage = g.STR_UI_PROCESS_XPERIENCE_SUBMISSION 
            strMessage += g.STR_UI_PROCESS_FAILED
            print(strMessage)

            # change the safety indicator
            blnSafety = False

        # discard the password
        del strPassword

    if pintChoice in lstScraping and strScrapeAbsence != 'c':
        # inform user about the process start
        print(g.STR_UI_PROCESS_ABSENCE_DOWNLOAD + g.STR_UI_PROCESS_STARTED)

        try:
            # scrape xperience data, process and save them to an external file
            blnSuccess = wps.blnObtainXperienceAbsences(
                strUserName,
                strPassword,
                strDateStart,
                strDateEnd,
                strScrapeAbsence
            )

            if blnSuccess:
                # inform the user about the process end if login didn't fail
                print(g.STR_UI_ABSENCE_SCRAPING)

        except:
            # process failed, inform the user
            print(g.STR_UI_PROCESS_ABSENCE_DOWNLOAD + g.STR_UI_PROCESS_FAILED)

            # change the safety indicator
            blnSafety = False

        # discard the password
        del strPassword

    if blnSafety and pintChoice in [
        g.INT_UI_CHOICE_OUTLOOK_IMPORT,
        g.INT_UI_CHOICE_FULL_DOWNLOAD
    ] and strScrapeAbsence != 'c':
        # inform user about the process start
        print(g.STR_UI_PROCESS_ABSENCE_SAVING + g.STR_UI_PROCESS_STARTED)

        try:
            # save xperience data to outlook
            cai.ImportAbsences()

            # inform the user about the process end
            print(g.STR_UI_ABSENCE_OUTLOOK)

        except:
            # process failed, inform the user
            print(g.STR_UI_PROCESS_ABSENCE_SAVING + g.STR_UI_PROCESS_FAILED)

    # the process ends here, say goodbye to the user
    print(g.STR_UI_GOODBYE)