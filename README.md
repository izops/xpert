# xper:t - submit your absence automatically

The code contains an interactive bot that analyzes your Outlook calendar and
submits your absences to Xperience attendance system. It also allows user to
scrape their own Xperience calendar and save the downloaded absences to
Outlook.


## Requirements

Installation of python with following modules required:
- maskpass `pip install maskpass` (only until version 2.1)
- pandas `pip install pandas`
- pywintypes `pip install pypiwin32`
- selenium `pip install selenium`
- win32com.client `pip install pywin32`


## Outlook convention

Depending on the version of xper:t there is a possibility to choose from a
convention of flagging a day as working from home/working from office. Please,
read the release notes to understand the convention.

Only the days containing all-day meeting flagged as 'working elsewhere', and 
the days containing no all-day meeting are considered for the submission to
Xperience system.


## Description

The xper:t bot uses Outlook API to read person's own calendar, extracts
information about all day meetings and outputs the findings to a text file.
It can also use the output of the calendar scraping and submit the information
to Xperience attendance system.  
 
Another option is to download submitted absences from Xperience and save them
in own Outlook calendar.

Note: As of now, xper:t can only submit home office absences to Xperience
system.


## Usage

Launch script called `launcher.py` and follow the instructions. The calendar
extraction provides data in the format required for the submission to
Xperience. Xperience downloader provides data in the format required for 
Outlook saving.

To use the latest version of the tool, use the version available on branch
'main'. To use a specific version of the tool, use the following command:
`git checkout <version number>`, for example `git checkout V1.1`
