# VBA code for Authentication and a History Log for tracking.
The VBA Code serves an excel document to allow authentication and history log tracking. 

## Authentication
Upon opening the excel sheet, the user is prompted to enter their userID which will be checked againsts the list under the worksheet 'AuthorizedUsers'. If there is not a match, access is denied. Upon successful access, the userID is saved and all subsequent edits made will be reflected in the history log, that the edit has been made by that specific user. The VBA Code checks the worksheet 'AuthorizedUsers' for the respective name of the user corresponding to the userID and reflects the name in the history log.

## History log
The history log tracks changes made in the worksheet 'Notebooks returned'. Any changes made will be reflected in the 'History log' worksheet in the VBA code. The timestamp (date,time) of the edit is also recorded next to the change details. Hence, the history log shows the 'Time stamp', 'Changed by:', 'Workbook' (worksheet where the change was made), 'Cell Address' (the cell where the change was made) and 'Changed to'.

## Misc
A Vlookup function is embedded in the VBA Code. When the Serial Number is entered into the column 'Serial No.' of the worksheet 'Notebooks returned', a vlookup of the worksheet 'Computer Details (For Vlookup)' populates the worksheet 'Notebooks returned' with the 'Computer Name', 'Model', 'Name' (of the owner of the computer). It also automatically logs the timestamp of the edit made.


