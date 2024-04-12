# UserOffboarding
User Offboarding Script for on-Prem Exchange/AD Environment

This project was initially a project written for an on-prem environment.  That environment has since moved to Hybrid 365.  Therefore, changes need to be made to the script to account for that.
Below is a list of changes the original script made followed by a list of changes that need to be made to that original script.

==================================================
CURRENT OFFBOARDING STEPS:
==================================================

<# --- Active Directory account dispensation section --- #>
#Pull a list of SSRS subscriptions and/or Watchdog Reports for the user and remove the user from those reports
#Remove/Retire user from TMW, TMT, and TotalMail
#Disable the account
#Generates a random 20-character password and converts it to plaintext for use in this script.
#Resets user's password
#set extensionAttribute10 to today's date for use when deleting the account
# Add the OU path where the account originally came from to the description of the account's properties
# Get the list of permissions (group names) and export them to a CSV file for safekeeping
# Strip the permissions from the account

<# --- Exchange email account dispensation section --- #>
#remove any previously configured forwarding rules
#configure forwarding to the Supervisor's email address
#set Out of Office on the user's mailbox.
#Export .pst file
#disable Exchange settings (OWA/ActiveSync/etc.)
# Move the account to the Disabled Users OU

==================================================
O365 OFFBOARDING STEPS:
==================================================

<# --- Active Directory account dispensation section --- #>
No change Needed: #Pull a list of SSRS subscriptions and/or Watchdog Reports for the user and remove the user from those reports
No change Needed: #Remove/Retire user from TMW, TMT, and TotalMail

Done using Microsoft Graph Powershell
#Disable the account
#Generates a random 20-character password and converts it to plaintext for use in this script.
#Resets user's password

No change Needed: #set extensionattribute10 to today's date for use when deleting the account

REMOVE: # Add the OU path where the account originally came from to the description of the account's properties

Done using Microsoft Graph Powershell
# Get the list of permissions (group names) and export them to a CSV file for safekeeping
# Strip the permissions from the account

<# --- Exchange email account dispensation section --- #>
No change Needed: #remove any previously configured forwarding rules

UPDATE #configure forwarding to the Supervisor's email address 
 Convert to Shared Mailbox and give the Supervisor full access instead

No change Needed: #set Out of Office on the user's mailbox.

UPDATE #Export .pst file - No longer needed?
	Write a script to export the Shared Mailbox before the account deletion script(or as part of the account deletion script)
      https://github.com/ruudmens/LazyAdmin/blob/master/Exchange/Export-Mailbox.ps1

No change Needed: #disable Exchange settings (OWA/ActiveSync/etc.)


REMOVE: # Move the account to the Disabled Users OU - Leave in place, updated Account deletion script?
Will adjust Account Deletion Script to search entire NAPA Users OU for accounts with extensionattribute10 -le 60 days

