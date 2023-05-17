# %%
# Contains code that launches process of Outlook calendar analysis and 
# submission of absences to Xperience system using selenium

# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import ui.general_ui as ugu
import ui.process_management as upm

# %% run code
# ask user which process to run
intSelection = ugu.intGreeting()

# run the selected process
upm.RunProcess(intSelection)