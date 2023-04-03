# Release notes

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