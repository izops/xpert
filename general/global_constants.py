# %% contains public constants for xperience automation
import logging

# set logging level
OBJ_LOGGING_LEVEL = logging.ERROR

# %% URLs
# login URL
STR_URL_LOGIN = 'https://my.xperience.app/zurich/login.jsp'

# add absence URL
STR_URL_ADD_ABSENCE = 'https://my.xperience.app/zurich/web/timesystem?'
STR_URL_ADD_ABSENCE += '__mvcevent=absenceList&id=0'

# list absences URL
STR_URL_LIST_ABSENCES = 'https://my.xperience.app/zurich/web/timesystem?__'
STR_URL_LIST_ABSENCES += 'mvcevent=absenceList'

# %% selenium elements - submission of absences
# IDs
STR_ELEMENT_ID_ABSENCE_NOTE = 'absenceDetailFM_note'
STR_ELEMENT_ID_ABSENCE_SUBMIT = 'absenceDetailFM_submitRequest'
STR_ELEMENT_ID_HOME = 'menu_item_home'

# xpaths
# absence type dropdown
STR_ELEMENT_XPATH_ABSENCE_TYPE = "//input[@class='ng-scope ng-isolate-scope"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " custom-combobox-input ui-widget ui-widget"
STR_ELEMENT_XPATH_ABSENCE_TYPE += "-content ui-state-default ui-corner-left"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " ui-autocomplete-input']"

# absences main menu
STR_ELEMENT_XPATH_ABSENCE_MAIN = "//div[@class='h1' and text() = 'Absencie']"

# absence detail
STR_ELEMENT_XPATH_ABSENCE_DETAIL = "//div[@class='h1' and text() = 'Detail"
STR_ELEMENT_XPATH_ABSENCE_DETAIL +=" absencie']"

# absence date
STR_ELEMENT_XPATH_ABSENCE_DATE = "//input[@class='textbox_date info flapps"
STR_ELEMENT_XPATH_ABSENCE_DATE += "-date aw-calendar hasDatepicker' and"

# absence start date
STR_ELEMENT_XPATH_ABSENCE_DATE_START = STR_ELEMENT_XPATH_ABSENCE_DATE 
STR_ELEMENT_XPATH_ABSENCE_DATE_START += " @name='startDate']"

# absence end date
STR_ELEMENT_XPATH_ABSENCE_DATE_END = STR_ELEMENT_XPATH_ABSENCE_DATE 
STR_ELEMENT_XPATH_ABSENCE_DATE_END += " @name='endDate']"

# absence error box
STR_ELEMENT_XPATH_ABSENCE_ERROR = "//div[@class = 'error']"

# %% selenium elements - download of submitted absences
# IDs
STR_ELEMENT_ID_DATE_FROM = 'absenceRequestListTM_fd_dateFrom'
STR_ELEMENT_ID_DATE_TO = 'absenceRequestListTM_fd_dateTo'
STR_ELEMENT_ID_CHECKBOX_OPEN = 'checkfia_status0'
STR_ELEMENT_ID_CHECKBOX_APPROVED = 'checkfia_status1'
STR_ELEMENT_ID_CHECKBOX_CLOSED = 'checkfia_status3'
STR_ELEMENT_ID_BUTTON_FILTER = 'absenceRequestListTM__find'


# xpaths
# filter
STR_ELEMENT_XPATH_STATUS = '//input[@class = "ui-widget ui-widget-content ui-'
STR_ELEMENT_XPATH_STATUS += 'corner-left ui-autocomplete-input"]'
STR_ELEMENT_XPATH_TYPE = '//input[@class = "ng-scope ng-isolate-scope custom-'
STR_ELEMENT_XPATH_TYPE += 'combobox-input ui-widget ui-widget-content ui-state'
STR_ELEMENT_XPATH_TYPE += '-default ui-corner-left ui-autocomplete-input"]'

# table
STR_ELEMENT_XPATH_TABLE = '//tbody[@class = "data"]'

# other
STR_ATTRIBUTE_CHECKED = 'checked'

# waits
FLT_MAX_LOGIN_TIMEOUT = 30.0

# %% website keywords
# absence keywords
STR_ABSENCE_TYPE_HOME_OFFICE = 'Home office'
STR_ABSENCE_TYPE_DOCTOR = 'Doctor'
STR_ABSENCE_TYPE_PERSONAL_DAY = 'Personal day'
STR_ABSENCE_TYPE_VACATION = 'Vacation'
STR_ABSENCE_TYPE_SICK_LEAVE = 'Sick leave (PN)'
STR_ABSENCE_TYPE_NONE = 'Working from office'

# scraping absences
LST_SCRAPE_ABSENCES = [
    STR_ABSENCE_TYPE_VACATION,
    STR_ABSENCE_TYPE_SICK_LEAVE,
    STR_ABSENCE_TYPE_PERSONAL_DAY
]

# %% regular expressions
# data input - date start, date end, absence type
STR_REGEX_DATA_INPUT = '((?:[1-9]|0[1-9]|[1-2][0-9]|3[0-1])\/(?:[1-9]|0[1-9]'
STR_REGEX_DATA_INPUT += '|1[0-2])\/20[0-9]{2})\t((?:[1-9]|0[1-9]|[1-2][0-9]|3'
STR_REGEX_DATA_INPUT += '[0-1])\/(?:[1-9]|0[1-9]|1[0-2])\/20[0-9]{2})\t(Home'
STR_REGEX_DATA_INPUT +=' office|Vacation|Personal day|Sick leave)'

# user input date
STR_REGEX_DATE = '^(20[0-9]{2})?(0[1-9]|1[0-2])?(0[1-9]|[1-2][0-9]|3[01])$'

# web scraped duration
STR_REGEX_ABSENCE_DURATION = '(\d+\.\d+)'

# %% Outlook constants
# API meeting status constants
INT_MEETING_FREE = 0
INT_MEETING_TENTATIVE = 1
INT_MEETING_BUSY = 2
INT_MEETING_OUT_OF_OFFICE = 3
INT_MEETING_WORKING_ELSEWHERE = 5

# all day status constant for mixed day absence (non-API)
INT_MEETING_MIXED = -1

# Office meeting status
STR_MEETING_FREE = 'Free'
STR_MEETING_BUSY = 'Busy'
STR_MEETING_TENTATIVE = 'Tentative'
STR_MEETING_OUT_OF_OFFICE = 'Out of office'
STR_MEETING_WORKING_ELSEWHERE = 'Working elsewhere'

STR_MEETING_MIXED = 'Mixed availability status'

# %% user interaction
# name
STR_UI_BOT_NAME = '''
                                                                      
                                                               ,d     
                                                               88     
8b,     ,d8  8b,dPPYba,    ,adPPYba,  8b,dPPYba,     888     MM88MMM  
 `Y8, ,8P'   88P'    "8a  a8P_____88  88P'   "Y8     888       88     
   >888<     88       d8  8PP"""""""  88                       88     
 ,d8" "8b,   88b,   ,a8"  "8b,   ,aa  88             888       88,    
8P'     `Y8  88`YbbdP"'    `"Ybbd8"'  88             888       "Y888  
             88                                                       
             88                                                       

'''

# intro message
STR_UI_INTRO = 'Hi%s, my name is xper:t and I can analyze your calendar, '
STR_UI_INTRO += 'submit your absences to Xperience (currently I do only '
STR_UI_INTRO += 'home office absences), or download your submitted absences '
STR_UI_INTRO += 'and save them to your Outlook calendar.\n'

# process options
STR_UI_OFFER = 'So, how can I help you%s? [1/2/3/4/5/c(ancel)]\n'
STR_UI_OFFER += '\t1. Analyze my calendar\n'
STR_UI_OFFER += '\t2. Submit my absences to Xperience\n'
STR_UI_OFFER += '\t3. Analyze my calendar and then submit my absences to '
STR_UI_OFFER += 'Xperience (1 + 2)\n'
STR_UI_OFFER += '\t4. Donwload submitted absences from Xperience\n'
STR_UI_OFFER += '\t5. Save downloaded absences to Outlook calendar\n'
STR_UI_OFFER += '\t6. Download submitted absences and save them in Outlook '
STR_UI_OFFER += '(4 + 5)\n'

LST_UI_ANSWERS_PROCESS = [
    '1', '2', '3', '4', '5', '6',
    'c', 'cancel', 'c(ancel)'
]

# process choices
INT_UI_CHOICE_OUTLOOK_EXPORT = 1
INT_UI_CHOICE_XPERIENCE_SUBMISSION = 2
INT_UI_CHOICE_FULL_SUBMISSION = 3
INT_UI_CHOICE_XPERIENCE_SCRAPE = 4
INT_UI_CHOICE_OUTLOOK_IMPORT = 5
INT_UI_CHOICE_FULL_DOWNLOAD = 6

# maximum relevant choices
INT_UI_CHOICES_MAX = INT_UI_CHOICE_FULL_DOWNLOAD

# request home office convention
STR_UI_REQUEST_CONVENTION = '\nPlease%s, tell me in what way do you use '
STR_UI_REQUEST_CONVENTION += '\'working elsewhere\' all day appointment:\n'
STR_UI_REQUEST_CONVENTION += '\t1. Working elsewhere means I work from home\n'
STR_UI_REQUEST_CONVENTION += '\t2. Working elsewhere means I work from '
STR_UI_REQUEST_CONVENTION += 'office\n'

LST_UI_ANSWERS_CONVENTION = ['1', '2']

# request date inputs
STR_UI_REQUEST_DATE = '\nPlease, provide <SELECT> date in YYYYMMDD format:\n'
STR_UI_REQUEST_DATE += '\tIf you want a date from current year, you can use '
STR_UI_REQUEST_DATE += 'MMDD only\n'
STR_UI_REQUEST_DATE += '\tIf you want a date from current month, you can use '
STR_UI_REQUEST_DATE += 'DD only\n'

STR_UI_REQUEST_PLACEHOLDER = '<SELECT>'
STR_UI_REQUEST_DATE_START = 'start'
STR_UI_REQUEST_DATE_END = 'end'

# process start info
STR_UI_PROCESS_STARTED = ' has started...\n'
STR_UI_PROCESS_CALENDAR_ANALYSIS = 'Calendar analysis'
STR_UI_PROCESS_XPERIENCE_SUBMISSION = 'Submission to Xperience'
STR_UI_PROCESS_ABSENCE_DOWNLOAD = 'Downloading of absences from Xperience'
STR_UI_PROCESS_ABSENCE_SAVING = 'Saving of absences to Outlook calendar'

# process fail info
STR_UI_PROCESS_FAILED = ' has failed unexpectedly. Please, contact the '
STR_UI_PROCESS_FAILED += 'maintainer of the repository for help.\n'

# login fail info
STR_UI_LOGIN_FAILED = 'Login failed. Okta login timed out without sign on.'

# goodbyes
STR_UI_GOODBYE_CANCEL = '\nAll right then%s, goodbye!'
STR_UI_CALENDAR_ANALYSIS_COMPLETE = '\n\nThe calendar analysis is complete.\n'
STR_UI_SUBMISSION_SUCCESS = '\nSubmission of absences finished '
STR_UI_SUBMISSION_SUCCESS += 'sucessfully.'
STR_UI_SUBMISSION_FAIL = '\nSubmission of absences failed. Check log for '
STR_UI_SUBMISSION_FAIL += 'more information.'
STR_UI_ABSENCE_SCRAPING = '\nAbsences were downloaded successfully.'
STR_UI_ABSENCE_OUTLOOK = '\nAbsences were saved in Outlook successfully.'
STR_UI_GOODBYE = '\nThanks for stopping by and have a nice day%s! (°͜°)/"'

# %% paths and file names
# calendar data path
STR_PATH_DATA = './data/'
STR_FILE_CALENDAR_DATA = 'calendar_data.txt'
STR_FULL_PATH_CALENDAR_DATA = STR_PATH_DATA + STR_FILE_CALENDAR_DATA

# log file name
STR_FILE_LOG = 'process.log'
STR_FULL_PATH_LOG = STR_PATH_DATA + STR_FILE_LOG

# scraped absences file output
STR_FILE_SCRAPED_DATA = 'absences_from_xperience.txt'
STR_FULL_PATH_SCRAPED_DATA = STR_PATH_DATA + STR_FILE_SCRAPED_DATA

# %% other constants
# user name
STR_USER_DOMAIN = '@zurich.com'

# calendar time period - hours
INT_CALENDAR_TIME_UNIT_HOURS = 60

# scraping data column names
LST_COLUMN_NAMES_SCRAPED = [
  'Name',
  'From',
  'To',
  'Duration',
  'Status',
  'Modified',
  'Absence_type'
]

# columns for scraped import
LST_COLUMN_NAMES_IMPORT = [
  'From',
  'To',
  'Duration',
  'Status',
  'Absence_type'
]

# default name of downloaded absence
STR_ABSENCE_CALENDAR_NAME = ': Xperience Download'

# global date format
STR_DATE_FORMAT = '%d/%m/%Y'