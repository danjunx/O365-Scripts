# Run this command to create a new Event Log for this script:
# New-EventLog -LogName application -Source "Office 365 Log"
# This script is EVENT ID 1

###################################################################################################################################################################
#							Cusomisable bits start here
###################################################################################################################################################################
# License SKU to assign to the accounts, use Get-MsolAccountSku to find current SKU's
$LicenseSKU = "TenantID:LICENSESKU"
# Usage location for the accounts, has to be the 2 digit country code
$UsageLocation = "GB"
# Add the event ID here to filter with in Event Viewer
$EventID = "1"

###################################################################################################################################################################
#							Office 365 email settings start here
###################################################################################################################################################################
$UserName = "O365SVCAcct@domain.tld"
$Password = cat "C:\Path\to\Keys\O365SVCAcct_domain.tld.key" | ConvertTo-SecureString

$SmtpServer = "smtp.office365.com"
$MailTo = @("it@domain.tld")
$MailFrom = "O365SVCAcct@domain.tld"
$MailPort = "587"
$MailSubject = "O365 Automation: License assignment"
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password

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
    write-eventlog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId '$EventID' -Message "AD/O365 AUTOMATION`n`nAssignO365Licenses`n`nError Connecting to Office 365! Please verify your user name and password"
    exit
}

catch {
    #Log for any other error
    Write-Host "Error Connecting"
    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId '$EventID' -Message "AD/O365 AUTOMATION`n`nAssignO365Licenses`n`nError Connecting to Office 365!"
    exit
}
###################################################################################################################################################################
#							Script starts here
###################################################################################################################################################################
#Get users that are on the domain and are unlicensed
$Users = Get-MsolUser -All -DomainName "domain.tld" |?{-not $_.IsLicensed}

$totalcount = $users.Count
# Get a total count of all the users

$HTMLBody = "
<head>
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
</head><body>
<table>
<tr><th>Name</th><th>License assigned</th><th>IsLicensed</th><th>Usage location</th><th>Correct License</th></tr>
"

if ($Users) {

    foreach ($User in $Users) {
        $UserUPN = $user.UserPrincipalName
        Set-MsolUser -UserPrincipalName $UserUPN -UsageLocation $UsageLocation
        Set-MsolUserLicense -UserPrincipalName $UserUPN -AddLicenses $LicenseSKU

        $GetUsageLocation = Get-MsolUser -UserPrincipalName $UserUPN | select UsageLocation | ft -HideTableHeaders | Out-String
        $GetLicense = Get-MsolUser -UserPrincipalName $UserUPN | select Licenses | ft -HideTableHeaders | Out-String
        $IsLicensed = Get-MsolUser -UserPrincipalName $UserUPN | select IsLicensed | ft -HideTableHeaders | Out-String

        if ($IsLicensed -like "*True*") {
            $IsLicensedCheck = '<font color="green">True</font>'
        } Else {
            $IsLicensedCheck = '<font color="red">False</font>'
        }

        if ($GetLicense -like "*$LicenseSKU*") {
            $CorrectLicense = '<font color="green">True</font>'
        } else {
            $CorrectLicense = '<font color="red">False</font>'
        }

        $HTMLBody += "<tr><td>$UserUPN</td><td>$LicenseSKU</td><td>$IsLicensedCheck</td><td>$GetUsageLocation</td><td>$CorrectLicense</td></tr>"
    }

    $HTMLBody += "</table>
    </body>"
    $SuccessMessage = "AD/O365 AUTOMATION`n`nSTUDENT LICENSE ASSIGNMENT AUTOMATION`n`nTotal Licenses assigned: $totalcount`nSee email to $MailTo for more details."

    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId $EventID -Message $SuccessMessage
    Send-MailMessage -To "$MailTo" -from "$MailFrom" -Subject $MailSubject -Body $HTMLBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials -BodyAsHtml
    #Write-Host "Users: $users"
    exit
}
else {
    Write-Host "Zero licenses were assigned"
    exit
}
