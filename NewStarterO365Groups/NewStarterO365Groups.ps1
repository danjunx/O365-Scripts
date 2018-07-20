# Run this command to create a new Event Log for this script:
# New-EventLog -LogName application -Source "Office 365 Log"
# This script is EVENT ID $EventID

###################################################################################################################################################################
#							Cusomisable bits start here
###################################################################################################################################################################
# Path to the XML that contains the new user details from the Microsoft Forms form
$XMLFilePath = "C:\path\to\OneDrive\Starter form responses\EmailGroupXMLFiles\XMLFile*.xml"
# Add the event ID here to filter with in Event Viewer
$EventID = 8
# This is an important variable that might change somewhere else
$UPNDomain = "domain.tld"
# List of email groups that should not be added to a new user, even if the SE has the groups
$EmailGroupBlacklist = @("AllStaffAllowed@domain.tld", "NewYorkAllowed@domain.tld", "LondonAllowed@domain.tld")
# Set the license usagelocation
$usagelocation = "GB"
# Set the license SKU that we want to apply to the user
$licensesku = "tennantname:STANDARDWOFFPACK_IW_FACULTY"

###################################################################################################################################################################
#							Office 365 email settings start here
###################################################################################################################################################################
# Use this part for anything related to O365 accounts or email settings (un/pw's/email subjects/etc)
$username = "O365SVCAcct@domain.tld"
$password = "password"

$SmtpServer = "smtp.office365.com"
$MailTo = @("helpdesk@domain.tld")
$MailFrom = "O365SVCAcct@domain.tld"
$MailPort = "587"
$MailSubjectSuccess = "AD Automation: Operation successful"
$MailSubjectFailed = "AD Automation: Operation failed"
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $($password | ConvertTo-SecureString -AsPlainText -Force)

###################################################################################################################################################################
#							Office 365 connection junk start here
###################################################################################################################################################################
try {
    #Attempts to connect to Office 365 and install Modules
    Import-Module MSOnline
    $pass = convertto-securestring -String "$password" -AsPlainText -Force
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $pass
    Connect-MsolService -Credential $credential -ErrorAction Stop
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
    Import-PSSession -AllowClobber $ExchangeSession >null
}
catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
    #Logs error for incorrect password
    #Write-Host "Please verify your username and password"
    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId $EventID -Message "AD/O365 AUTOMATION`n`nNEW STARTER O365 GROUPS AUTOMATION`n`nError Connecting to Office 365! Please verify your user name and password"
    exit
}

catch {
    #Log for any other error
    #Write-Host "Error Connecting"
    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId $EventID -Message "AD/O365 AUTOMATION`n`nNEW STARTER O365 GROUPS AUTOMATION`n`nError Connecting to Office 365!"
    exit
}

###################################################################################################################################################################
#							Script starts here
###################################################################################################################################################################
# Test to see if the XML exists, if it does continue
$XMLPath = get-childitem $XMLFilePath -recurse
$TodaysDate = Get-Date -UFormat "%d/%m/%Y"

if (test-path $XMLPath) {
    foreach ($XML in $XMLPath) {
        # Try and replace each ampersand inside the script with an escaped ampersand
        (get-content $xml).replace('&', '&amp;') | set-content $xml

        [XML]$NewUser = Get-Content $XML
        $NewName = $($NewUser.User.NewUser)
        $SecurityEquivalent = $($NewUser.User.SecurityEquivalent)

        write-host "--------------------------------------------------------------------------"
        write-host "New user: $NewName"
        write-host "Security Equivalent user: $SecurityEquivalent"
        write-host "--------------------------------------------------------------------------"

        $SEUser = Get-User $SecurityEquivalent
        $SEGroups = Get-Group | Where-Object {$_.Members -contains $SEUser}

        Set-MsolUser -UserPrincipalName "$NewName@$UPNDomain" -UsageLocation $usagelocation
        write-host "Set $NewName's usage location to $UsageLocation."
        Set-MsolUserLicense -UserPrincipalName "$NewName@$UPNDomain" -AddLicenses "$licensesku"
        write-host "Added license $licensesku to $NewName's product licenses."
        Set-Mailbox "$NewName@$UPNDomain" -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Update, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create, UpdateFolderPermission -AuditDelegate Update, SoftDelete, HardDelete, SendAs, Create, UpdateFolderPermissions, MoveToDeletedItems, SendOnBehalf -AuditOwner UpdateFolderPermission, MailboxLogin, Create, SoftDelete, HardDelete, Update, MoveToDeletedItems

        foreach ($SEGroup in $SEGroups) {
            if ($SEGroup.WindowsEmailAddress -in $EmailGroupBlacklist) {
                write-host "Skipping group $SEGroup because it's in the group blacklist."
            } else {

                if ($SEGroup.RecipientTypeDetails -eq "MailUniversalDistributionGroup") {
                    Add-DistributionGroupMember -Identity $SEGroup.WindowsEmailAddress -Member "$NewName@$UPNDomain"
                    write-host "Added distribution group member $SEGroup."

                } ElseIf ($SEGroup.RecipientTypeDetails -eq "GroupMailbox") {
                Add-UnifiedGroupLinks -Identity $SEGroup.WindowsEmailAddress -LinkType Members -Links "$NewName@$UPNDomain"
                write-host "Added unified group links $SEGroup."

                } Else {
                    write-host "Skipping $SEGroup because it's not either a Unified group or Distribution list."
                }
            }
        }

        remove-item $XML
        if (test-path $XML) {
            write-host "XML still exists"
        } else {
            write-host "XML has been deleted"
        }
    }
} else {
# If the XML doesn't exist, don't do anything
write-host "XML does not exist, nothing to do"
}
