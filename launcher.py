# %%
# Contains code that launches process of Outlook calendar analysis and 
# submission of absences to Xperience system using selenium

# %% import modules
import sys

sys.path.append('../emea_oth_xpert')
import ui.general_ui as gi
import ui.absences_process as ap

# %% run code
# ask user which process to run
intSelection = gi.intGreeting()

# run the selected process
ap.RunProcess(intSelection)