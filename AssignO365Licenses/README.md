# AssignO365Licenses
This script checks your O365 tenant for unlicensed users on a specific domain, and assigns licenses to them. <br/> In detail, the script will find users on your tenant without a license on a specific domain, if the user doesn't have a license, it will  first set the usage location for the user, and then assign a license to the user determined by a variable inside the script. Once the license has been "assigned", the script will check both the users usage location and their licenses to confirm it was successful. Once it's finished it will generate a HTML formatted table with the results for each user.

### Prerequisites/things to note
- None this time!

### How to use
In our environment, the setup for this script is as follows:
- Generic O365 "service" account that is an "Exchange administrator" and "User management administrator" on O365.
- Generic AD "service" account for the script to run as.
- Scripts run from the same server as the Azure AD Connect agent (useful but not necessary). This is so we can trigger the script to run after an ADConnect sync has finished running.

To get things started, you'll want to download and save the script somewhere where it can run uninterrupted, then edit and import the included scheduled task with your details (Account to run the script as, location of the vbs file, etc)<br/>

After doing this, you'll want to run through the top part of the script and edit the variables (usage location, license SKU, account details, etc)<br/>
A quick easy way to get your license SKU's is to run a ``Get-MsolAccountSku`` after connecting to O365.<br/>
To check licenses on an existing account, run a ``Get-MsolUser -UserPrincipalName "user@domain.tld" | select Licenses``.

I've included the vbs file to launch the script from (without it popping up a blank PowerShell window every time), a scheduled task XML file that you should be able to edit with the filepath/username and then import into Task Scheduler. <br />
In our case, we have a scheduled task that runs when event ID 114 is triggered for the source Directory Synchronization (A successful AAD sync).

## Other useful info
- Every time I update the script in our production environment I'll try to update this repo
- I've probably missed some bits off of this readme, but if I do notice anything I'll add it to This
