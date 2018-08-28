# Run this command to create a new Event Log for this script:
# New-EventLog -LogName application -Source "Office 365 Log"
# This script is EVENT ID $EventID

###################################################################################################################################################################
#							Cusomisable bits start here
###################################################################################################################################################################
# Path to the XML that contains the new user details from the Microsoft Forms form
$XMLFilePath = "C:\path\to\OneDrive\Starter form responses\Starter form response*.xml"
# Path to the location where generated XML files are stored for email group assignment
$XMLFilePathEmail = "C:\path\to\OneDrive\Starter form responses\EmailGroupXMLFiles\"
# Path for failed XML files to live in without being deleted by the script (If the account exists already, etc)
$XMLFilePathFailed = "C:\path\to\OneDrive\Starter form responses\EmailGroupXMLFiles\FailedResponses\"
# Add the event ID here to filter with in Event Viewer
$EventID = 5
# Add the UPN domain here, make sure it's the full domain.tld
$UPNDomain = "domain.tld"
# Set a plaintext defualt password for the account here, has to meet complexity requirements
#$DefaultPasswordPT = "Password1!"
# REPLACED, see line 94
# Add the IT DG email address
$MailToIT = "it@domain.tld"
# Add AD groups to blacklist here, make sure it's the full CN,CN,DC,DC and that they're in "quotes" within the array
$ADGroupBlacklist = @("CN=Administrators,CN=Builtin,DC=domain,DC=tld", "CN=Domain Admins,CN=Users,DC=domain,DC=tld")
# Add users to whitelist here, make sure it's the full first.last@domain.tld and that they're in "quotes" in the array
$FormWhitelist = @("user.name@domain.tld","user1@domain.tld","user2@domain.tld")

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
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $password
$FormCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $FormUsername, $FormPassword
###################################################################################################################################################################
#							Script starts here
###################################################################################################################################################################

Import-Module ActiveDirectory
Import-Module NameIT
# Test to see if the XML exists, if it does continue
$XMLPath = get-childitem $XMLFilePath -recurse
$TodaysDate = Get-Date -UFormat "%d/%m/%Y"

if (test-path $XMLPath) {
    foreach ($XML in $XMLPath) {
        # Try and replace each ampersand inside the script with an escaped ampersand
        (get-content $xml).replace('&', '&amp;') | set-content $xml

        [XML]$NewUser = Get-Content $XML
        $NewName = $($NewUser.NewUser.Name)
        $StartDate = $($NewUser.NewUser.StartDate)
        $LineManager = $($NewUser.NewUser.LineManager)
        $Department = $($NewUser.NewUser.Department)
        $Site = $($NewUser.NewUser.Site)
        $JobTitle = $($NewUser.NewUser.JobTitle)
        $OfficeEquipment = $($NewUser.NewUser.OfficeEquipment)
        $SecurityEquivalent = $($NewUser.NewUser.SecurityEquivalent)
        $Notes = $($NewUser.NewUser.Notes)
        $SubmittedBy = $($NewUser.NewUser.SubmittedBy)
        $FormTodaysDate =  $($NewUser.NewUser.TodaysDate)

        if ($SubmittedBy -in $FormWhitelist) {

            if ($FormTodaysDate -eq $TodaysDate) {

                # Debug messages
                $NewFormMessage = "AD/O365 AUTOMATION`n`nNEW STAFF ACCOUNT CREATION`nXML exists, new form has probably been submitted"
                #Send-MailMessage -To "$MailTo" -from "$MailFrom" -Subject $MailSubjectSubmitted -Body $NewFormMessage -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials
                Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $NewFormMessage
                Write-host "XML exists, new form has probably been submitted"
                # Split the names given from the XML into seperate variables
                $FirstNewName,$LastNewName = $NewName.Split(' ')
                $SEFirstName,$SELastName = $SecurityEquivalent.Split(' ')
                $ManagerFirstName,$ManagerLastName = $LineManager.Split(' ')
                # Rejoin the names back into a single variable with a dot (.) inbetween them to make the username
                $NewNameJoined = "$FirstNewName.$LastNewName"
                $SENameJoined = "$SEFirstName.$SELastName"
                $LineManagerJoined = "$ManagerFirstName.$ManagerLastName"
                # Set the default password variable here rather than at the top of the script (generates different passwords for each form)
                $DefaultPasswordPT = Invoke-Generate [adjective][noun]### -ApprovedVerb
                # Convert the default password to a securestring variable to use later
                $DefaultPasswordSS = $DefaultPasswordPT | ConvertTo-SecureString -AsPlainText -Force
                # Create the email address to use later
                $NewEmailAddress = "$NewNameJoined@$UPNDomain"


                # Set $CreateInOU and $HomeDirectory variables depending on what site is specified in the XML
                if ($Site -like "London") {
                    $CreateInOU = "OU=London,OU=Company users,DC=domain,DC=tld"
                    $HomeDirectory = "\\LON-File01\UserHome\$NewNameJoined"
                    #write-host "CreateInOU: $CreateInOU `nHomeDirectory: $HomeDirectory"
                } ElseIf ($Site -like "New York") {
                    $CreateInOU = "OU=New York,OU=Company users,DC=domain,DC=tld"
                    $HomeDirectory = "\\NYC-File01\UserHome\$NewNameJoined"
                    #write-host "CreateInOU: $CreateInOU `nHomeDirectory: $HomeDirectory"
                } ElseIf ($Site -like "Amsterdam") {
                    $CreateInOU = "OU=Amsterdam,OU=Company users,DC=domain,DC=tld"
                    $HomeDirectory = "\\AMS-File01\UserHome\$NewNameJoined"
                    #write-host "CreateInOU: $CreateInOU `nHomeDirectory: $HomeDirectory"
                } Else {
                    $FailMessageSite = "AD/O365 AUTOMATION`n`nNEW STAFF ACCOUNT CREATION`nInvalid site specified, not creating account."
                    Send-MailMessage -To "$MailTo" -from "$MailFrom" -Subject $MailSubjectFailed -Body $FailMessageSite -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials
                    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId $EventID -Message $SuccessMessage
                    #write-host "$FailMessageSite"
                    exit
                }
                # Check to see if the user account already exists within AD, if it doesn't, continue
                $CheckNewNameExists = get-aduser -filter {sAMAccountName -eq $NewNameJoined}
                   if (!$CheckNewNameExists) {
                       New-ADUser -Name "$NewName" -GivenName "$FirstNewName" -Surname "$LastNewName" -SamAccountName "$NewNameJoined" -UserPrincipalName "$NewNameJoined@$UPNDomain" -EmailAddress "$NewNameJoined@$UPNDomain" -AccountPassword $DefaultPasswordSS -Path "$CreateInOU" -HomeDirectory "$HomeDirectory" -HomeDrive "H:" -Title "$JobTitle" -Manager "$LineManagerJoined" -Department "$Department" -Enabled $True
                       # Get the info of the new account we just created so we can grab the SID later
                       $NewNameJoinedSID = get-aduser -Identity $NewNameJoined
                       # Check to see if the home directory already exists or not
                        if (test-path $HomeDirectory) {
                        write-host "HomeDirectory "$HomeDirectory" already exists"
                        } else {
                            # Create the new home folder, use -force so it doesn't ask us for an input y/n
                            new-item -path $HomeDirectory -type directory -force

                            $acl = Get-Acl $HomeDirectory

                            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ($NewNameJoinedSID.SID, "FullControl", "ContainerInherit,ObjectInherit", "none", "Allow")
                            $acl.AddAccessRule($AccessRule)
                            Set-Acl -Path $HomeDirectory -AclObject $acl
                        }

                        Get-ADUser $NewNameJoined -Properties MailNickName | Set-ADUser -Replace @{MailNickName = "$FirstNewName $LastNewName"}

                        # Get the security groups of the Security Equivalent user and copy them to the new user we just created
                       $SEGroupsFriendly = Get-ADPrincipalGroupMembership $SENameJoined | select name | Format-Table -HideTableHeaders | Out-String
                       $SEGroups = Get-ADUser -Identity $SENameJoined -Properties MemberOf | Select-Object -ExpandProperty memberof
                       ForEach ($Group in $SEGroups) {
                           # If the group is one specified in the $ADGroupBlacklist variable at the top of the script, skip it and write-host to confirm skipped, else continue adding the groups to the new user
                           if ($Group -in $ADGroupBlacklist) {
                               write-host "Skipping blacklisted group: $Group"
                            } Else {
                                #Write-Host "Add-ADGroupMember -Identity $Group -Member $NewNameJoined"
                                Add-ADGroupMember -Identity $Group -Member $NewNameJoined
                           }
                       }
                       write-host "AD account created"


                       # Generate an XML file to handle the email group assignment
                       # Set the file name, putting it here ensures it'll be a random name every time the foreach runs
                       $XMLFileNameEmail = Invoke-Generate "XMLFile#####.xml"
                       # Create the file using the path variable at the top of the script
                       $XmlWriter = New-Object System.XMl.XmlTextWriter("$XMLFilePathEmail$XMLFileNameEmail",$Null)
                       # Set the formatting
                       $xmlWriter.Formatting = "Indented"
                       $xmlWriter.Indentation = "4"
                       # Write to the file
                       $xmlWriter.WriteStartElement("User")
                       $xmlWriter.WriteElementString("NewUser","$NewNameJoined")
                       $xmlWriter.WriteElementString("SecurityEquivalent","$SENameJoined")
                       $xmlWriter.WriteEndElement
                       $xmlWriter.WriteEndElement()
                       # Finish and close the document
                       $xmlWriter.Finalize
                       $xmlWriter.Flush
                       $xmlWriter.Close()

                    } else {
                        # If the user already exists, send an email to report the problem and move the file elsewhere so the script doesn't send emails continuously
                       $FailMessageAlreadyExists = "AD/O365 AUTOMATION`n`nNEW STAFF ACCOUNT CREATION`n`User $FirstNewName $LastNewName already exists in AD, not creating account.`nMoved XML file to $XMLFilePathFailed.`nNew user form submitted by $SubmittedBy."
                       Send-MailMessage -To "$MailToIT" -from "$MailFrom" -Subject $MailSubjectFailed -Body $FailMessageAlreadyExists -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials
                       Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId $EventID -Message $FailMessageAlreadyExists
                       write-host "AD account exists already"
                       move-item -Path $XML -Destination $XMLFilePathFailed
                       exit
                   }

                   # Delete the XML once completed
                remove-item $XML
                if (test-path $XML) {
                    write-host "XML still exists"
                } else {
                    write-host "XML has been deleted"
                    # Create a string with the new user account details and write to the event log
                    $SuccessMessage = "AD/O365 AUTOMATION
                    `n`nNEW STAFF ACCOUNT CREATION
                    `n`nStaff Name: $NewName
                    `nUsername: $NewNameJoined
                    `nDefault password: $DefaultPasswordPT
                    `nEmail address: $NewEmailAddress
                    `n`nCreated in OU: $CreateInOU
                    `n`nMember of the following Security Groups: $SEGroupsFriendly
                    `nNew user form submitted by: $SubmittedBy"
                    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $SuccessMessage
                    #send-MailMessage -To "$MailTo" -from "$MailFrom" -Subject $MailSubjectCreated -Body $SuccessMessage -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials
                    # Create the email to send to the ticket system and send it off
                    $TicketEmailSubject = "NEW STARTER - $NewName"
                    $TicketEmailBody = "Date starting: $StartDate
                    `nName: $NewName
                    `nLine manager: $LineManager
                    `nDepartment: $Department
                    `nSite: $site
                    `nJob title: $JobTitle
                    `nOffice Equipment Required: $OfficeEquipment
                    `nSecurity equivalent to...: $SecurityEquivalent
                    `nNotes: $Notes
                    `nUsername: $NewNameJoined
                    `nDefault password: $DefaultPasswordPT
                    `nCreated in OU: $CreateInOU
                    `nMember of the following Security Groups: $SEGroupsFriendly
                    `nNew user form submitted by: $SubmittedBy"
                    send-MailMessage -To "$MailTo" -from "$FormUsername" -Subject $TicketEmailSubject -Body $TicketEmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials
                    # Create the email to send to HR/the person that submitted the form and send it off
                    $HREmailSubject = "New Starter - $NewName"
                    $HREmailBody = "Name: $NewName
                    `nJob title: $JobTitle
                    `nDate starting: $StartDate
                    `nLine manager: $LineManager
                    `nDepartment: $Department
                    `nSite: $site
                    `nOffice Equipment Required: $OfficeEquipment
                    `nSecurity equivalent to...: $SecurityEquivalent
                    `nNotes: $Notes"
                    send-MailMessage -To "$SubmittedBy" -from "$FormUsername" -Subject $HREmailSubject -Body $HREmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials

                }
            }
        } else {
            #Form has been submitted by someone not in the whitelist at the top of the script
            $NotFromEmailSubject = "New Starter Form - Unauthorised form submitted"
            $NotFromEmailBody = "New starter form has been submitted by an unauthorised users
            `nSubmitted by: $SubmittedBy
            `nName: $NewName
            `nDepartment: $Department
            `nNew user creation process skipped for this form"
            send-MailMessage -To "$MailToIT" -from "$FormUsername" -Subject $NotFormEmaiLSubject -Body $NotFormEmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $FormCredentials
            Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $NotFormEmailBody
            Remove-Item $XML
        }
    }
    # Try and run an AD sync after the account has been created in AD, just to speed the process up a little bit
    Start-ADSyncSyncCycle -PolicyType Delta
} else {
# If the XML doesn't exist, don't do anything
    write-host "XML does not exist, nothing to do"
}
