# Search-GUI Tool

## Authors:  
Agustin Gallegos  
Nelson Riera  

## Version History:
### 3.40 - 05/11/2020
 - Updated tool to connect to Exchange Online using new EXO v2 module.
### 3.30 - 01/17/2018
 - Added Radio button for simplicity when running app with no existing PS Session
### 3.20 - 01/17/2018
 - Fixed missing "Import-PSSession" import commands missing. thanks to "diego.a.sanchez" for reporting it
### 3.10 - 01/11/2018
 - Added 2 additional buttons. "Import from CSV", in order to import a list of users to work with
 - "Generate Log only", in order to generate a full list of search results, to a target mailbox
### 3.00 - 01/08/2018
 - Added Get-RecoverableItems and Restore-RecoverableItems capability
### 2.40 - 06/06/2017
 - Included a "From" and "To" fields to be combined in search filters
### 2.30 - 06/02/2017
 - Included "Is Soft-Deleted" Checkbox in order to search in a soft-deleted account
### 2.20 - 03/13/2017
 - Determined "premise" variable is not populated if there is an existing PSSession, fixing it
 - Determined "cred" variable is not populated if there is an existing PSSession, hence breaking Permissions validation
 - Changed permissions validation to an additional button
 - Added cosmetic status bar to Main Window
### 2.00 - 01/08/2017
 - Added "All Available Mailboxes" Checkbox in order to work on all Mailboxes. Credits to colleague Ramon Rocha in LATAM Escalation team for this suggestion
 - Change resultant output to DataGrid. This allows to view results in a grid, and alows to copy cells values for reference
 - Added Time logging in Powershell window for the operations
 - Rolled back changes for the "Direction" ComboBox. Couldn't find the option to have a default selection
 - Added cosmetic link to "Subject" line. This links to website with reference to KQL syntax
### 1.01 - 01/05/2017
 - Corrected "Direction" ComboBox with a default selection and not leaving blank
### 1.00 - 01/05/2017
 - First release
### 1.00 - 12/23/2016
 - Project start