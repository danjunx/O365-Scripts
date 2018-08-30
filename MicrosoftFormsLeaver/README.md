# MicrosoftFormsLeaver
This script takes XML responses generated from a Microsoft Forms/Flow workflow and disables the corresponding AD/O365 accounts. <br/> In detail, the script will disable the AD account, reset the password, remove the user from the O365 GAL, disable the sign-in status and apps (IMAP, POP3, etc), set an OoO, and set email forwarding to another user. It will also check against our asset management system [SnipeIT](https://snipeitapp.com) for any assets assigned to the user.<br/> Once the script has done the thing it sends an email to the user that submitted the form with what they submitted, and sends an email to our helpdesk with all the details for the user it just disabled. <br/> The script can work without Forms/Flow, you'll just need to generate the XML file yourself somehow.

### Prerequisites/things to note
- This script uses [NameIT](https://github.com/dfinke/NameIT) to randomly generate the password string.
- This script uses [SnipeitPS](https://github.com/snazy2000/SnipeitPS) to query SnipeIT and get the users assigned assets.
- Our SnipeIT installation has LDAP Integration enabled and assets are assigned to AD users (instead of internal SnipeIT users).
- In our production environment, this script runs on the same server as our Azure AD Connect agent, so we can trigger an ADConnect sync at the end of the script.

### How to use
In our environment, the setup for this script is as follows:
- Generic O365 "forms" account that the Forms and Flows live and run on.
- Generic O365 "service" account that is an "Exchange administrator" and "User management administrator" on O365.
- Scripts run from the same server as the Azure AD Connect agent(useful but not necessary).
- OneDrive for Business is installed on the server and logged in as the "forms" account (So we can access the XML files that the Flow creates).
- Generic AD service account that's used to create/modify the AD users.

To get things started, you'll want to create a Form on the generic "forms" account that has all the fields required for a leaver (usually the full name, when their last day onsite is, who their emails are going to forwarded to, who is going to have their data, etc)<br/>
Like so: <br/>
<img src="/Images/MicrosoftFormsLeaver/MicrosoftForm.png" width="60%">

Once you have your form sorted, you'll need to create a Flow to save new Forms responses to a folder in Onedrive folder. Here I'm using the submission time as part of the filename so we can have multiple forms submitted at once without it causing problems.<br/> The File Content is generated as an XML-like file using Dynamic Content and the responses within the form to fill the fields. One thing to note is that inserting the Dynamic Content variables on the Flow creation page was a bit fiddly, but that may have just been a problem at my end.  
<img src="/Images/MicrosoftFormsLeaver/MicrosoftFlow.png" width="60%">

If you're using the SnipeIT part of the script, you'll want to do the following:
- Login to SnipeIT as a generic account that can be used to pull data(only needs to be able to read, not put or modify).
- Head over to Admin > Manage API Keys.
- Create New Token, and call it whatever you like (I called mine MicrosoftFormsPowershellScript) and hit Create.
- Copy the token out and paste it into the $SnipeITAPIKey variable inside the script.
- While you're inside, you can also set the $SnipeITURL variable.

After doing this, the next thing to do is put the script somewhere it can run without interruptions (the server AADConnect is running is a good place), install OneDrive for Business on the server, and read/edit the top section of the script with usernames/passwords/filepaths. (There's also some useful comments in there related to event logging)<br />

I've included an example of the XML file the Flow generates, a vbs file to launch the script from (without it popping up a blank PowerShell window every time), and a scheduled task XML file that you should be able to edit with the filepath/username and then import into Task Scheduler. <br />
In our case, we have a scheduled task running every day at 6PM as the Generic AD Service account.

## Other useful info
- Every time I update the script in our production environment I'll try to update this repo
- I've probably missed some bits off of this readme, but if I do notice anything I'll add it to This
