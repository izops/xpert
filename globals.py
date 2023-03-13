# %% contains public constants for xperience automation

# %% URLs
# login URL
STR_URL_LOGIN = 'https://my.xperience.app/zurich/login.jsp'

# add absence URL
STR_URL_ADD_ABSENCE = 'https://my.xperience.app/zurich/web/timesystem?'
STR_URL_ADD_ABSENCE += '__mvcevent=absenceList&id=0'

# %% selenium elements - IDs
STR_ELEMENT_ID_USERNAME = 'user'
STR_ELEMENT_ID_PASSWORD = 'pwd'
STR_ELEMENT_ID_LOGIN = 'loginButton'
STR_ELEMENT_ID_LOGIN_ERROR = 'errorBox'
STR_ELEMENT_ID_ABSENCE_NOTE = 'absenceDetailFM_note'
STR_ELEMENT_ID_ABSENCE_SUBMIT = 'absenceDetailFM_submitRequest'

# %% selenium elements - xpaths
# absence type dropdown
STR_ELEMENT_XPATH_ABSENCE_TYPE = "//input[@class='ng-scope ng-isolate-scope"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " custom-combobox-input ui-widget ui-widget"
STR_ELEMENT_XPATH_ABSENCE_TYPE += "-content ui-state-default ui-corner-left"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " ui-autocomplete-input']"

# absences main menu
STR_ELEMENT_XPATH_ABSENCE_MAIN = "//div[@class='h1' and text() = 'Absencie'"

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

# %% website keywords
# absence keywords
STR_ABSENCE_TYPE_HOME_OFFICE = 'home office'
STR_ABSENCE_TYPE_DOCTOR = 'doctor'
STR_ABSENCE_TYPE_PERSONAL_DAY = 'personal day'
STR_ABSENCE_TYPE_VACATION = 'vacation'
STR_ABSENCE_TYPE_SICK_LEAVE = 'sick leave'
STR_ABSENCE_TYPE_NONE = 'working from office'

# %% regular expressions
# data input - date start, date end, absence type
STR_REGEX_DATA_INPUT = '((?:[1-9]|0[1-9]|[1-2][0-9]|3[0-1])\/(?:[1-9]|0[1-9]'
STR_REGEX_DATA_INPUT += '|1[0-2])\/20[0-9]{2})\t((?:[1-9]|0[1-9]|[1-2][0-9]|3'
STR_REGEX_DATA_INPUT += '[0-1])\/(?:[1-9]|0[1-9]|1[0-2])\/20[0-9]{2})\t(Home'
STR_REGEX_DATA_INPUT +=' office|Vacation|Personal day|Sick leave)'

# %% Outlook constants
# API meeting status constants
INT_MEETING_FREE = 0
INT_MEETING_TENTATIVE = 1
INT_MEETING_BUSY = 2
INT_MEETING_OUT_OF_OFFICE = 3
INT_MEETING_WORKING_ELSEWHERE = 5

# Office meeting status
STR_MEETING_FREE = 'Free'
STR_MEETING_BUSY = 'Busy'
STR_MEETING_TENTATIVE = 'Tentative'
STR_MEETING_OUT_OF_OFFICE = 'Out of office'
STR_MEETING_WORKING_ELSEWHERE = 'Working elsewhere'


# %% other constants
# user name
STR_USER_DOMAIN = '@zurich.com'

# calendar data path
STR_PATH_CALENDAR_DATA = 'c:/repositories/emea_oth_xpert/data/'
STR_FILE_CALENDAR_DATA = 'calendar_data.txt'
STR_FULL_PATH_CALENDAR_DATA = STR_PATH_CALENDAR_DATA + STR_FILE_CALENDAR_DATA

# calendar time period - hours
INT_CALENDAR_TIME_UNIT_HOURS = 60