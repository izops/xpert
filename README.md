# xper:t - submit your absence automatically

The code contains an interactive bot that analyzes your Outlook calendar and
submits your absences to Xperience attendance system.


## Requirements

Installation of python with following modules required:
- win32com.client `pip install pywin32`
- pywintypes `pip install pypiwin32`
- selenium `pip install selenium`


## Outlook convention

- Every day that *does not* contain all-day meeting is considered to be
**working from home**
- Every day that contains all-day meeting flagged as **working elsewhere** in
Outlook is considered to be **working from office**
- Every day that contains all-day meeting flagged as **out of office** or
**busy** is excluded from the submission to Xperience
- If a day contains **out of office** or **working elsewhere** flagged
meeting(s) for only part of the day, is excluded from the submission to
Xperience


## Description

The xper:t bot uses Outlook API to read person's own calendar, extracts
information about all day meetings and outputs the findings to a text file.
It can also use the output of the calendar scraping and submit the information
to Xperience attendance system.

Note: As of now, xper:t can only submit home office absences to Xperience
system.


## Usage

Launch script called `launcher.py` and follow the instructions. The calendar
extraction provides data in the format required for the submission to Xperience.

To use the latest version of the tool, use the version available on branch
'main'. To use a specific version of the tool, use the following command:
`git checkout <version number>`


## Contact

To report a problem, please contact
[the author](mailto:ivan.zustiak@zurich.com).