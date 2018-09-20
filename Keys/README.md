# Keys folder

Use this folder to store keys for use inside of the PowerShell Scripts

To generate the keys, use the following command: <br />
``Read-Host -Prompt “Enter your account password” -AsSecureString | ConvertFrom-SecureString | Out-File "C:\Path\To\Keys\username_domain.tld.key"``

For these keys to work properly, you'll need to run PowerShell as the user your scripts are running as, on the server your scripts are running from. <br /> In my case, this was as simple as shift right-clicking PowerShell and selecting "Run as differnent user" and feeding the prompt with the service account details.
