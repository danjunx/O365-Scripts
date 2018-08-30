# Run this command to create a new Event Log for this script:
# New-EventLog -LogName application -Source "Office 365 Log"
# This script is EVENT ID $EventID

# This script uses SnipeitPS and NameIT
# https://github.com/snazy2000/SnipeitPS
# https://github.com/dfinke/NameIT

###################################################################################################################################################################
#							Cusomisable bits start here
###################################################################################################################################################################
# Path to the XML that contains the new user details from the Microsoft Forms form
$XMLFilePath = "C:\path\to\OneDrive\Leaver form responses\Leaver form response*.xml"
# Add the event ID here to filter with in Event Viewer
$EventID = 7
# Add the backup location for the users data, must be a valid path and the user the script runs as must have write permissions
$LeaverBackupPath = "\\LON-FILE01\LeaverProfiles\"
# Path to disabled users OU
$DisabledUsersOU = "OU=Pending,OU=Disabled Users,OU=Company users,DC=domain,DC=tld"
# Set the email address domain
$UPNDomain = "domain.tld"
# Company number field
$CompanyNumber = "01234 567890"
# Company name field
$CompanyName = "Comtech Corp LTD"
# Add the IT DG email address
$MailToIT = "it@domain.tld"
# Add users to whitelist here, make sure it's the full first.last@domain.tld and that they're in "quotes" in the array
$FormWhitelist = @("user.name@domain.tld","user1@domain.tld","user2@domain.tld")
# Set the URL for Snipeit
$SnipeITURL = "http://LON-SNIPEIT.DOMAIN.TLD"
# Set the API key for SnipeIT
$SnipeITAPIKey = "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"
###################################################################################################################################################################
#							Office 365 email settings start here
###################################################################################################################################################################
$username = "O365SVCAcct@domain.tld"
$Password = cat "C:\Path\to\Keys\O365SVCAcct_domain.tld.key" | ConvertTo-SecureString
$FormUsername = "O365FormsAcct@domain.tld"
$FormPassword = cat "C:\Path\to\Keys\O365FormsAcct_domain.tld.key" | ConvertTo-SecureString

$SmtpServer = "smtp.office365.com"
$MailTo = @("helpdesk@domain.tld")
$MailFrom = "O365SVCAcct@domain.tld"
$MailPort = "587"
$MailSubjectSubmitted = "AD Automation: New user form submitted"
$MailSubjectCreated = "AD Automation: New user created"
$MailSubjectFailed = "AD Automation: New user creation failed"
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $Password
$FormCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $FormUsername, $FormPassword

###################################################################################################################################################################
#							Office 365 connection junk starts here
###################################################################################################################################################################
try {
    #Attempts to connect to Office 365 and install Modules
    Import-Module MSOnline
    Connect-MsolService -Credential $Credentials -ErrorAction Stop
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credentials -Authentication "Basic" -AllowRedirection
    Import-PSSession -AllowClobber $ExchangeSession >null
}
catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
    #Logs error for incorrect password
    Write-Host "Please verify your username and password"
    write-eventlog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId '$EventID' -Message "AD/O365 AUTOMATION`n`nMicrosoftFormsLeaver`n`nError Connecting to Office 365! Please verify your user name and password"
    exit
}

catch {
    #Log for any other error
    Write-Host "Error Connecting"
    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId '$EventID' -Message "AD/O365 AUTOMATION`n`nMicrosoftFormsLeaver`n`nError Connecting to Office 365!"
    exit
}
###################################################################################################################################################################
#							Script starts here
###################################################################################################################################################################

Import-Module ActiveDirectory
Import-Module SnipeitPS
Import-Module NameIT

Set-Info -URL $SnipeITURL -apiKey $SnipeITAPIKey
# Test to see if the XML exists, if it does continue
$XMLPath = get-childitem $XMLFilePath -recurse
$TodaysDate = Get-Date -UFormat "%Y-%m-%d"

if (test-path $XMLPath) {
    foreach ($XML in $XMLPath) {

        [XML]$Leaver = Get-Content $XML
        $LeaverName = $($Leaver.Leaver.Name)
        $LeavingDate = $($Leaver.Leaver.LeavingDate)
        $DataGivenTo = $($Leaver.Leaver.DataGivenTo)
        $EmailGivenTo = $($Leaver.Leaver.EmaiLGivenTo)
        $EmailForwardedTo = $($Leaver.Leaver.EmailForwardedTo)
        $Notes = $($Leaver.Leaver.Notes)
        $SubmittedBy = $($Leaver.Leaver.SubmittedBy)

        if ($LeavingDate -le $TodaysDate) {

            if ($SubmittedBy -in $FormWhitelist) {

                write-host "Todays leaver is $LeaverName"
                # Debug messages
                $NewFormMessage = "AD/O365 AUTOMATION`n`nLEAVER PROCESS`nXML exists, new form has probably been submitted`n$LeaverName"
                Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $NewFormMessage
                Write-host "XML exists, new form has probably been submitted"
                # Split the names given from the XML into seperate variables
                $LeaverFirstName,$LeaverLastName = $LeaverName.Split(' ')
                $DGTFirstName,$DGTLastName = $DataGivenTo.Split(' ')
                $EGTFirstName,$EGTLastName = $EmailGivenTo.Split(' ')
                $EFTFirstName,$EFTLastName = $EmailForwardedTo.Split(' ')
                # Rejoin the names back into a single variable with a dot (.) inbetween them to make the username
                $LeaverNameJoined = "$LeaverFirstName.$LeaverLastName"
                $DGTNameJoined = "$DGTFirstName.$DGTLastName"
                $EGTNameJoined = "$EGTFirstName.$EGTLastName"
                $EFTNameJoined = "$EFTFirstName.$EFTLastName"
                # Create the email address to use later
                $LeaverEmailAddress = "$LeaverNameJoined@$UPNDomain"

                Write-host $LeaverNameJoined


                # Check to see if the user account exists within AD, if it does, continue
                $CheckLeaverExists = get-aduser -filter {sAMAccountName -eq $LeaverNameJoined}
                if (!$CheckLeaverExists) {
                    # If the leaver doesn't exist, send an email to report the problem
                    $FailMessageDoesNotExist = "AD/O365 AUTOMATION`n`nLEAVER PROCESS`n`User $LeaverName doesn't exist in AD, cannot continue with leaver process."
                    Send-MailMessage -To "$MailtTo" -from "$MailFrom" -Subject $MailSubjectFailed -Body $FailMessageDoesNotExist -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials
                    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId $EventID -Message $FailMessageDoesNotExist
                    write-host "AD account doesn't exist"
                    exit

                } else {
                    $LeaverDN = get-aduser -identity $LeaverNameJoined | select DistinguishedName | Format-Table -HideTableHeaders | Out-String
                    # Reset the leavers password to break O365
                    # Generate random password
                    $DefaultPasswordPT = Invoke-Generate [adjective][noun]### -ApprovedVerb
                    # Convert the default password to a securestring variable to use later
                    $DefaultPasswordSS = $DefaultPasswordPT | ConvertTo-SecureString -AsPlainText -Force
                    # Set the AD password
                    Set-ADAccountPassword -Identity $LeaverDN.trim() -NewPassword $DefaultPasswordSS
                    # Disable the leavers AD account
                    Disable-ADAccount -Identity $LeaverDN.trim()
                    # Hide the user from the GAL
                    Set-aduser $LeaverDN.trim() -Replace @{msExchHideFromAddressLists=$true}
                    # Set the description of the user to the date Disabled
                    Set-ADUser $LeaverDN.Trim() -Description "Account disabled on $TodaysDate by leavers form/script"
                    # Get their DN so we can move them to a differnet OU
                    Move-ADobject $LeaverDN.trim() -TargetPath $DisabledUsersOU
                    # Disable sign-in on the O365 portal
                    Set-MsolUser -UserPrincipalName $LeaverEmailAddress -BlockCredential $true
                    # Disable email apps (OWA, MAPI, EWS, EAS, IMAP and POP)
                    Set-CASMailbox $LeaverEmailAddress -OWAEnabled $False -MAPIEnabled $False -EWSEnabled $False -ActiveSyncEnabled $False -ImapEnabled $False -PopEnabled $False
                    # Set forward to the EFT value
                    set-Mailbox $LeaverEmailAddress -ForwardingAddress "$EFTNameJoined@$UPNDomain"
                    Set-MailboxAutoReplyConfiguration -Identity $LeaverEmailAddress -AutoReplyState Enabled -InternalMessage "Please note that I am no longer working for $CompanyName, my responsibilities have been passed to $EFTFirstName $EFTLastName. Please contact he/she on $EFTNameJoined@$UPNDomain or $CompanyNumber." -ExternalMessage "Please note that I am no longer working for $CompanyName, my responsibilities have been passed to $EFTFirstName $EFTLastName. Please contact he/she on $EFTNameJoined@$UPNDomain or $CompanyNumber."
                }

                #SnipeIT parts
                $UserAsset = Get-Asset | Where-Object {$_.assigned_to -like "*$LeaverNameJoined*" }
                $HTMLAsset = "Asset/s<br>
                -----------------------------------<br>"

                foreach ($Asset in $UserAsset) {
                    # Select variable from the list of multiple variables it returns back
                    $AssetName = $Asset.name
                    $AssetTag = $Asset.asset_tag
                    $AssetSerial = $Asset.serial
                    $AssetManufacturer = $Asset.manufacturer | select name | Format-Table -HideTableHeaders | Out-String
                    $AssetManufacturerTrim = $AssetManufacturer.trim()

                    $AssetModel = $Asset.model | select name | Format-Table -HideTableHeaders | Out-String
                    $AssetModelTrim = $AssetModel.trim()

                    $AssetCheckout = $Asset.last_checkout | select formatted | Format-Table -HideTableHeaders | Out-String
                    $AssetCheckoutTrim = $AssetCheckout.trim()

                    $HTMLAsset += "Asset name: $AssetName<br>
                    Asset tag: $AssetTag<br>
                    Asset serial: $AssetSerial<br>
                    Asset manufacturer: $AssetManufacturerTrim<br>
                    Asset model: $AssetModelTrim<br>
                    Last checkout date: $AssetCheckoutTrim<br>
                    -----------------------------------<br>"

                }
                # Delete the XML once completed
                remove-item $XMLPath
                if (test-path $XMLPath) {
                    write-host "XML still exists"
                } else {
                    write-host "XML has been deleted"
                    # Create a string with the new user account details and write to the event log
                    $SuccessMessage = "AD/O365 AUTOMATION
                    `n`nLEAVER PROCESS
                    `n`nStaff Name: $LeaverName
                    `nUsername: $LeaverNameJoined
                    `nChanged password: $DefaultPasswordPT
                    `nEmail address: $LeaverEmailAddress
                    `nMoved to OU: $DisabledUsersOU
                    `nLeaver form submitted by: $SubmittedBy"
                    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $SuccessMessage
                    ##send-MailMessage -To "$MailTo" -from "$MailFrom" -Subject $MailSubjectCreated -Body $SuccessMessage -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials

                    # Create the email to send to the ticket system and send it off
                    $TicketEmailSubject = "LEAVER - $LeaverName"
                    $TicketEmailBody = "<font face='Calibri' color=#000000>Date leaving: $LeavingDate<br>
                    Name: $LeaverName<br>
                    Username: $LeaverNameJoined<br>
                    Email address: $LeaverEmailAddress<br>
                    Changed password: $DefaultPasswordPT<br>
                    Data to be given to: $DataGivenTo<br>
                    Email backup to be given to: $EmailGivenTo<br>
                    Email has been forwarded to: $EmailForwardedTo<br>
                    Moved to OU: $DisabledUsersOU<br>
                    Notes: $Notes<br>
                    -----------------------------------<br>
                    $HTMLAsset
                    Leaver form submitted by: $SubmittedBy<br>
                    </font>"

                    send-MailMessage -To "$MailTo" -from "$FormUsername" -Subject $TicketEmailSubject -Body $TicketEmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials -BodyAsHtml

                    # Create the email to send to HR/the person that submitted the form and send it off
                    $HREmailSubject = "Leaver - $LeaverName"
                    $HREmailBody = "<font face='Calibri' color=#000000>Name: $LeaverName<br>
                    Date leaving: $LeavingDate<br>
                    Data to be given to: $DataGivenTo<br>
                    Email backup to be given to: $EmailGivenTo<br>
                    Email to be forwarded to: $EmailForwardedTo<br>
                    Notes: $Notes<br></font>"
                    send-MailMessage -To "$SubmittedBy" -from "$FormUsername" -Subject $HREmailSubject -Body $HREmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials -BodyAsHtml

                    # Try and run an AD sync after the account has been created in AD, just to speed the process up a little bit
                    Start-ADSyncSyncCycle -PolicyType Delta
                }

            } else {
                #Form has been submitted by someone not in the whitelist at the top of the script
                $NotFromEmailSubject = "Leaver Form - Unauthorised form submitted"
                $NotFromEmailBody = "<font face='Calibri' color=#FF0000>A leaver form has been submitted by an unauthorised user<br><font>
                <font face='Calibri' color=#000000>Submitted by: $SubmittedBy<br>
                Leaver name: $LeaverName<br><br>
                New user creation process skipped for this form<br>"
                send-MailMessage -To "$MailToIT" -from "$FormUsername" -Subject $NotFromEmailSubject -Body $NotFromEmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials -BodyAsHtml
                Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $NotFormEmailBody
                Remove-Item $XML
            }
        } else {
            write-host "Leaver is not leaving until $LeavingDate"
        }
    }

} else {
    # If the XML doesn't exist, don't do anything
    write-host "XML does not exist, nothing to do"
}
