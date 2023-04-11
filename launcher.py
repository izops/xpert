# %%
# Contains code that launches process of Outlook calendar analysis and 
# submission of absences to Xperience system using selenium

# %% import modules
import ui as u

# %% run code
# ask user which process to run
intSelection = u.intGreeting()

# run the selected process
u.RunProcess(intSelection)