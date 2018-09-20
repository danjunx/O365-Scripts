command = "powershell.exe -nologo -command C:\path\to\script\AssignO365Licenses.ps1"
 set shell = CreateObject("WScript.Shell")
 shell.Run command,0
