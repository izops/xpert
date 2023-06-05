# %%
# Contains code that launches process of Outlook calendar analysis and 
# submission of absences to Xperience system using selenium

# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import ui.general_ui as ugu
import ui.process_management as upm
import calendar_works.secondary_data as csd

# %% run code
# obtain user name for personalization
tplName = csd.tplGetFullUserName()

# ask user which process to run
intSelection = ugu.intGreeting(tplName)

# run the selected process
upm.RunProcess(intSelection, tplName)