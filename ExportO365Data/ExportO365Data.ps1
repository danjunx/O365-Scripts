# Run this command to create a new Event Log for this script:
# New-EventLog -LogName application -Source "Office 365 Log"
# This script is EVENT ID $EventID

###################################################################################################################################################################
#							Cusomisable bits start here
###################################################################################################################################################################
# Add the event ID here to filter with in Event Viewer
$EventID = 11
# Filepath for the XML files to pull the data from
$XMLFilePath = "C:\path\to\OneDrive\Leaver form responses\O365DataExport\ExportO365Data*.xml"
# Local export location to act as a "buffer" before data is copied off to the network share
$LocalExportLocation = "C:\Path\to\O365Exports"
# Location for the export either \\server.domain.tld\share$ or c:\folder, cannot have trailing backslash
$RemoteExportLocation = "\\LON-FILE01\LeaverProfiles"
# Path to microsoft.office.client.discovery.unifiedexporttool.exe. Usually found in %LOCALAPPDATA%\Apps\2.0\
$UnifiedExportTool = "C:\path\to\UnifiedExportTool\microsoft.office.client.discovery.unifiedexporttool.exe"
# Description for the export to show in O365
$ExportDescription = "ExportO365Data auto-generated search"
# Add the IT DG email address
$MailToIT = "it@domain.tld"

###################################################################################################################################################################
#							Office 365 email settings start here
###################################################################################################################################################################
# Use this part for anything related to O365 accounts or email settings (un/pw's/email subjects/etc)
$UserName = "O365SVCAcct@domain.tld"
$Password = cat "C:\path\to\Keys\O365SVCAcct_domain.tld.key" | ConvertTo-SecureString

$SmtpServer = "smtp.office365.com"
$MailTo = @("IT@domain.tld")
$MailFrom = "o365svcacct@domain.tld"
$MailPort = "587"
$MailSubjectSuccess = "AD Automation: Operation successful"
$MailSubjectFailed = "AD Automation: Operation failed"
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $Password

###################################################################################################################################################################
#							Office 365 connection junk starts here
###################################################################################################################################################################
try {
    #Attempts to connect to Office 365 and install Modules
    Import-Module MSOnline
    Connect-MsolService -Credential $Credentials -ErrorAction Stop
    $ComplianceSearch = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" -Credential $Credentials -Authentication Basic -AllowRedirection
    Import-PSSession -AllowClobber $ComplianceSearch >null
}
catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
    #Logs error for incorrect password
    Write-Host "Please verify your username and password"
    write-eventlog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nError Connecting to Office 365! Please verify your user name and password"
    exit
}

catch {
    #Log for any other error
    Write-Host "Error Connecting"
    Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nError Connecting to Office 365!"
    exit
}
###################################################################################################################################################################
#							Script starts here
###################################################################################################################################################################

$XMLPath = get-childitem $XMLFilePath -recurse
$TodaysDate = Get-Date -UFormat "%d/%m/%Y"

if (test-path $XMLFilePath) {
    foreach ($XML in $XMLPath) {
        [XML]$Export = Get-Content $XML
        $ExportUser = $($Export.User.ExportUser)
        $ExportUserName = $($Export.User.ExportUserName)
        write-host "--------------------------------------------------------------------------"
        write-host "Export user: $ExportUser"
        write-host "Export username: $ExportUserName"
        write-host "--------------------------------------------------------------------------"
        Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance search pending for user $ExportUser"

        $SearchName = "$ExportUser-Search"

        New-ComplianceSearch -Name $SearchName -ExchangeLocation $ExportUserName -Description $ExportDescription

        Start-ComplianceSearch -Identity $SearchName
        do{
            Start-Sleep -s 5
            $ComplianceSearch = Get-ComplianceSearch $SearchName
            Write-Host "Compliance Search waiting..."
        }
        while ($ComplianceSearch.Status -ne 'Completed')
        Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance search completed for user $ExportUser"

        $SearchNameExport = "$ExportUser-Search_Export"
        New-ComplianceSearchAction -SearchName $SearchName -ExchangeArchiveFormat PerUserPst -Format FxStream -Export
        do{
            Start-Sleep -s 5
            $SearchComplete = Get-ComplianceSearchAction -Identity "$SearchNameExport"
            Write-Host "Compliance Search Action waiting..."
        }
        while ($SearchComplete.Status -ne 'Completed')
        Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance search action completed for user $ExportUser"

$ExportTemplate = @'
Container url: {ContainerURL*:https://*.blob.core.windows.net/*}; SAS token: {SASToken:?sv=2014-02-14&sr=*&si=eDiscoveryBlobPolicy*&sig=*%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 2; Exchange item format: Msg; Exchange archive format: IndividualMessage; SharePoint archive format: SingleZip; Include SharePoint versions: True; Enable dedupe: EnableDedupe:True; Reference action: "<null>"; Region: ; Started sources: StartedSources:3; Succeeded sources: SucceededSources:1; Failed sources: 0; Total estimated bytes: 12,791,334,934; Total estimated items: 143,729; Total transferred bytes: {TotalTransferredBytes:7,706,378,435}; Total transferred items: {TotalTransferredItems:71,412}; Progress: {Progress:49.69 %}; Completed time: ; Duration: 00:50:43.9321895; Export status: {ExportStatus:DistributionCompleted}
Container url: {ContainerURL*:https://*.blob.core.windows.net/*}; SAS token: {SASToken:?sv=2014-02-14&sr=*&si=eDiscoveryBlobPolicy*&sig=*%3D}; Scenario: General; Scope: BothIndexedAndUnindexedItems; Scope details: AllUnindexed; Max unindexed size: 0; File type exclusions for unindexed: <null>; Total sources: 1; Exchange item format: FxStream; Exchange archive format: PerUserPst; SharePoint archive format: IndividualMessage; Include SharePoint versions: True; Enable dedupe: True; Reference action: "<null>"; Region: ; Started sources: 2; Succeeded sources: 2; Failed sources: 0; Total estimated bytes: 69,952,559,461; Total estimated items: 107,707; Total transferred bytes: {TotalTransferredBytes:70,847,990,489}; Total transferred items: {TotalTransferredItems:100,808}; Progress: {Progress:93.59 %}; Completed time: 4/27/2018 11:45:46 PM; Duration: 04:31:21.1593737; Export status: {ExportStatus:Completed}
'@

        $ExportName = $SearchName + "_Export"
        $ExportDetails = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details | select -ExpandProperty Results | ConvertFrom-String -TemplateContent $exporttemplate
        $ExportDetails
        $ExportContainerURL = $ExportDetails.ContainerURL
        $ExportSASToken = $ExportDetails.SASToken

        # Download the exported files from Office 365
        Write-Host "Initiating download"
        Write-Host "Saving export to: $LocalExportLocation"
        $Arguments = "-name ""$SearchName""","-source ""$ExportContainerURL""","-key ""$ExportSASToken""","-dest ""$LocalExportLocation""","-trace true"
        Start-Process -FilePath "$UnifiedExportTool" -ArgumentList $Arguments -wait
        # Script will hang here while the tool is exporting

        # Wait 15 seconds for the UnifiedExportTool to cleanup after itself
        Write-Host "Waiting for 15 seconds for UnifiedExportTool to cleanup"
        Start-Sleep -s 15

        $CheckExportFinished = (Test-Path "$LocalExportLocation\$SearchName\Temp")
        if ($CheckExportFinished) {
            write-host "Export failed"
            Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance export failed for user $ExportUser"
            send-MailMessage -To "$MailToIT" -from "$MailFrom" -Subject "LEAVER - $ExportUser" -Body "Compliance export failed for user $ExportUser" -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials -BodyAsHtml

            exit
        } else {
            write-host "Export completed"
            Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance export completed for user $ExportUser"
            Move-Item -Path "$LocalExportLocation\$SearchName\*\Exchange" -Destination "$RemoteExportLocation\$ExportUser\Exchange"
            Move-Item -Path "$LocalExportLocation\$SearchName\*" -Destination "$RemoteExportLocation\$ExportUser\Exchange"
        }

        $CheckExportMoved = (Test-Path "$RemoteExportLocation\$ExportUser\Exchange\*.pst")
        if (!$CheckExportMoved) {
            write-host "Export move failed"
            Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Error -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance export move to $RemoteExportLocation failed for $ExportUser"
            send-MailMessage -To "$MailToIT" -from "$MailFrom" -Subject "LEAVER - $ExportUser" -Body "Compliance export move failed for user $ExportUser" -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials -BodyAsHtml

            exit
        } else {
            write-host "Export move completed"
            Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "AD/O365 AUTOMATION`n`nExportO365Data`n`nCompliance search moved to $RemoteExportLocation\$ExportUser"

            remove-item -Path "$LocalExportLocation\$SearchName" -Recurse

            remove-item $XML
            if (test-path $XML) {
                write-host "XML still exists"
            } else {
                write-host "XML has been deleted"
                $ExportCompleteEmailSubject = "Leaver - $ExportUser"
                $ExportCompleteEmailBody = "<font face='Calibri' color=#000000>Name: $ExportUser<br>
                O365 email export complete<br>
                Export saved to <a href='$RemoteExportLocation\$ExportUser'>$RemoteExportLocation\$ExportUser</a><br>
                </font>"

                send-MailMessage -To "$MailToIT" -from "$MailFrom" -Subject $ExportCompleteEmailSubject -Body $ExportCompleteEmailBody -SmtpServer $SmtpServer -port $MailPort -UseSsl -Credential $Credentials -BodyAsHtml
            }

        }

    }
} else {
# If the XML doesn't exist, don't do anything
    write-host "XML does not exist, nothing to do"
    #Write-EventLog -LogName Application -Source "Office 365 Log" -EntryType Information -EventId "$EventID" -Message "Nothing to do here..."
}
