# %% contains public constants for xperience automation

# URLs
# login URL
STR_URL_LOGIN = 'https://my.xperience.app/zurich/login.jsp'

# add absence URL
STR_URL_ADD_ABSENCE = 'https://my.xperience.app/zurich/web/timesystem?'
STR_URL_ADD_ABSENCE += '__mvcevent=absenceList&id=0'

# user name
STR_USER_DOMAIN = '@zurich.com'

# selenium elements
# IDs
STR_ELEMENT_ID_USERNAME = 'user'
STR_ELEMENT_ID_PASSWORD = 'pwd'
STR_ELEMENT_ID_LOGIN = 'loginButton'
STR_ELEMENT_ID_LOGIN_ERROR = 'errorBox'

# xpaths
STR_ELEMENT_XPATH_ABSENCE_TYPE = "//input[@class='ng-scope ng-isolate-scope"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " custom-combobox-input ui-widget ui-widget"
STR_ELEMENT_XPATH_ABSENCE_TYPE += "-content ui-state-default ui-corner-left"
STR_ELEMENT_XPATH_ABSENCE_TYPE += " ui-autocomplete-input']"

STR_ELEMENT_XPATH_ABSENCE_DETAIL = "//div[@class='h1' and text() = 'Detail"
STR_ELEMENT_XPATH_ABSENCE_DETAIL +=" absencie']"

# absence keywords
STR_ABSENCE_TYPE_HOME_OFFICE = 'home office'
STR_ABSENCE_TYPE_DOCTOR = 'doctor'
STR_ABSENCE_TYPE_PERSONAL_DAY = 'personal day'
STR_ABSENCE_TYPE_VACATION = 'vacation'
STR_ABSENCE_TYPE_SICK_LEAVE = 'sick leave'