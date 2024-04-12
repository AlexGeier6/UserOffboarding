
$date = [datetime]::Today.ToString('dd-MM-yyyy')
$todaysDate = get-date -Format 'MM-dd-yyy'
# Un-comment the following if PowerShell isn't already set up to do this on its own
Import-Module ActiveDirectory
Import-Module sqlserver


 #Blank the console
 Clear-Host
Write-host "Enter the SQL Server Credentials" -BackgroundColor White
$credential = Get-Credential
$serverName = 'SQLSVR01'
try {
        Invoke-Sqlcmd -ServerInstance $serverName -Database TMW_Live -Credential $credential -query "SELECT GETDATE() AS TimeOfQuery;"
        
}
catch {
        
        Start-Sleep -Seconds 3
        Exit
}

#If user passes checks, the script continues
Write-Host "Offboard a user"

<# --- Active Directory account dispensation section --- #>
$sam = Read-Host 'Account name to disable'

# Get the properties of the account and set variables
$user = Get-ADuser $sam -properties canonicalName, distinguishedName, displayName, mailNickname, mail
$dn = $user.distinguishedName
$cn = $user.canonicalName
$din = $user.displayName
$offboardedEmail = $user.mail

$UserManager = (Get-ADUser (Get-ADUser $sam -Properties manager).manager -Properties mail).mail
$AutoReply = "I am no longer with *Company*. If you need assistance please reach out to " + $UserManager + "."

#Offboarding Logs for memberOf, SSRS, and DawgAlerts
$Query1 = 'remove_email_from_dawg_and_ssrs_subscriptions'
$path1 = "\\fileserver\Offboarding logs\"
$path2 = "-AD-DisabledUserPermissions.csv"
$pathFinal = $path1 + $din + $path2
$SQLPath1 = "\\fileserver\Offboarding logs\"
$SQLPath2 = "-AD-DisabledUserPermissions.txt"
$SQLpathFinal = $SQLPath1 + $din + $SQLPath2


Try {
        
        #Pull a list of SSRS subscritions and/or Watchdog Reports for the user on remove the user from those reports
        $query1Results = Invoke-Sqlcmd -ServerInstance $serverName -Database TMW_Live -Credential $credential -query "[dbo].[remove_email_from_dawg_and_ssrs_subscriptions] @offboarded_email = '$offboardedEmail'"  
        $query1Results | Out-File -FilePath $SQLpathFinal
        $SQLresultsHTML = $query1Results | ConvertTo-Html -Property 'report_name' | Out-String

        #Remove/Retire user from TMW, TMT, and TotalMail
        Invoke-Sqlcmd -ServerInstance $serverName -Database TMW_Live -Credential $credential -query "[dbo].[remove_tmw_and_totalmail_login] @email_address = '$offboardedEmail'"  

        # Disable the account
        Disable-ADAccount $sam
        Write-Host ($din + "'s Active Directory account is disabled.")

        #Generates a random 20 character password and converts it to plaintext for use in this script.
        $Passwd = -join ((48..122) | Get-Random -Count 20 | ForEach-Object{[char]$_})
        $PasswdSecStr = ConvertTo-SecureString $passwd -AsPlainText -Force

        #Resets user's password
        Set-ADAccountPassword -Identity "$sam" -NewPassword $PasswdSecStr -Reset
        Write-Host ($din + "'s Active Directory password has been changed.")

        #set extensionAttribute to todays date for use when deleting the account
        Set-ADUser -Identity "$sam" -Clear "extensionAttribute10"
        Set-ADUser -Identity "$sam" -Add @{extensionAttribute10= "$todaysDate"}

        # Add the OU path where the account originally came from to the description of the account's properties
        Set-ADUser $dn -Description ("Moved from: " + $cn + " - on $date")
        Write-Host ($din + "'s Active Directory account path saved.")

Start-Sleep -Seconds 3

        # Get the list of permissions (group names) and export them to a CSV file for safekeeping
        $groupinfo = get-aduser $sam -Properties memberof | Select-Object name, 
        @{ n="GroupMembership"; e={($_.memberof | ForEach-Object{get-adgroup $_}).name}}
        $count = 0
        $arrlist =  New-Object System.Collections.ArrayList
    do{
        $null = $arrlist.add([PSCustomObject]@{
        # Name = $groupinfo.name
        GroupMembership = $groupinfo.GroupMembership[$count]
        })
        $count++ 
    }until($count -eq $groupinfo.GroupMembership.count)

        $arrlist | Select-Object groupmembership |
        convertto-csv -NoTypeInformation |
        Select-Object -Skip 1 |
        out-file $pathFinal
        Write-Host ($din + "'s Active Directory group memberships (permissions) exported and saved to " + $pathFinal)

        # Strip the permissions from the account
        Get-ADUser $User -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object{Remove-ADGroupMember $_ -member $User -Confirm:$false}
        Write-Host ($din + "'s Active Directory group memberships (permissions) stripped from account")

<# --- Exchange email account dispensation section --- #>
# Import the Exchange snapin (assumes desktop PowerShell)
        if (!(Get-PSSnapin | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.SnapIn"})) { 
	    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://MAILSERVER.local/Powershell -Authentication Kerberos
        Import-PSSession $Session -DisableNameChecking -AllowClobber
              
        #remove any previously configured forwarding rules
        Set-Mailbox -Identity "$sam" -forwardingsmtpaddress $null
        Set-Mailbox -Identity "$sam" -forwardingaddress $null

        #configure forwarding to Supervisor's email address
        Set-Mailbox -Identity "$sam" -forwardingsmtpaddress  $UserManager -DeliverToMailboxAndForward $true

        #set Out of Office on the user's mailbox.
        Set-MailboxAutoReplyConfiguration -Identity "$sam" -AutoReplyState Enabled -InternalMessage $AutoReply -ExternalMessage $AutoReply

        # Loop flag variables
        $Go1 = 0
        $Go2 = 0
        $Go3 = 0
        $GoDone = 0

       Function Save-File ([string]$initialDirectory) {

	    $PresAdmin = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	    $AdminCheck = Get-ManagementRoleAssignment -RoleAssignee "$PresAdmin" -Role "Mailbox Import Export" -RoleAssigneeType user
	    If ($AdminCheck -eq $Null) {New-ManagementRoleAssignment -Role "Mailbox Import Export" -User $PresAdmin}

	    $MailBackupFileDate = (get-date -UFormat %b-%d-%Y_%I.%M.%S%p)
	    $MailBackupInitialPath = "\\oldemployeeemailpst\"
	    $MailBackupFileName = $sam+$MailBackupFileDate+".pst"

        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms
    
        $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $OpenFileDialog.initialDirectory = $MailBackupInitialPath
        $OpenFileDialog.filter = "PST (*.pst)| *.pst"
	$OpenFileDialog.FileName = $MailBackupFileName
        $OpenFileDialog.ShowDialog() | Out-Null

        return $OpenFileDialog.filename

}

        #Export .pst file
        $MailBackupFile = Save-File
        New-MailboxExportRequest -Mailbox $sam -FilePath $MailBackupFile

        #disable Exchange settings (OWA/ActiveSync/etc.)
        Set-CasMailbox -Identity "$sam" -OWAEnabled $false -ActiveSyncEnabled $false -PopEnabled $false -ImapEnabled $false -OWAforDevicesEnabled $False
        
        # Move the account to the Disabled Users OU
        Move-ADObject -Identity $dn -TargetPath "Ou=Terminated,OU=Users,DC=napa,DC=local"
        Write-Host ($din + "'s Active Directory account moved to 'Terminated' OU")

        }

$To         = ($UserManager)
$From       = 'IT@company.com'
$SmtpServer = 'mail.company.com'
$Subject    = ($din + ' was successfully offboarded') 
$Body       = @"
<p>The following changes have been made to the user's account:<br />
Active Directory account is disabled.<br />
The User's email has been forwarded to their Manager.<br />
An automatic reply has been enabled of the user's mailbox.<br />
The password has been changed.<br />
Account path saved.<br />
Group memberships (permissions) exported and saved to '\\fileserver\Offboarding logs\'<br />
Group memberships (permissions) were stripped from the account.<br />
The account moved to Terminated OU<br />
Mailbox .pst was exported and saved to drive.<br />
Exchange settings were disabled (ActiveSync/OWA/etc.).<br />
The user has been removed from the following Reports:<br />
<p $SQLresultsHTML </p>
"@

Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SmtpServer
}
Catch
{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Send-MailMessage -From 'IT@company.com' -To 'alex.geier@company.com' -Subject "EmployeeOffboarding Script has failed to disable a user account" -SmtpServer 'mail.company.com' -Body "The error message is: '$ErrorMessage' $FailedItem"
    Break
}
