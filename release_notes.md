# Release notes
--------------------------------------------------------------------------------
## V1.2

Author: Ivan Zustiak  
Date: 11 Apr 2023  
Description:  
- User can choose from the convention - either 'working elsewhere' is considered
to be a home office absence, or it is considered to be a work from office

Outlook convention:  
- User can choose which is flagged as working elsewhere - home office or work
from office
- Out of office (full day or partial) appointments are not considered for
Xperience submission, partial working elsewhere is not considered either
--------------------------------------------------------------------------------
## V1.1

Author: Ivan Zustiak  
Date: 3 Apr 2023  
Description:  
- Calendar reference issue when turning to daylight saving time fixed
- Reads error messages from Xperience when submitting the absences and logs them
into the console, and to an external text file
- Process runs through all days in the requested time period and writes a log
about the success or failure of the absence submission  
  
Outlook convention:  
The convention is the same as in `V1.0`
--------------------------------------------------------------------------------
## V1.0

Author: Ivan Zustiak  
Date: 23 Mar 2023  
Description:  
- Calendar analysis on full-day basis, with full day meeting status return
- Home office absence identified as a calendar entry without any all day
meetings and without any out-of-office/working elsewhere meetings
- Submission of home office absences to xperience
- Process stops when a submission error occurs  
  
Outlook convention:  
- Every day that *does not* contain all-day meeting is considered to be
**working from home**
- Every day that contains all-day meeting flagged as **working elsewhere** in
Outlook is considered to be **working from office**
- Every day that contains all-day meeting flagged as **out of office** or
**busy** is excluded from the submission to Xperience
- If a day contains **out of office** or **working elsewhere** flagged
meeting(s) for only part of the day, is excluded from the submission to
Xperience
