command = "powershell.exe -nologo -command C:\path\to\script\ExportO365Data.ps1"
 set shell = CreateObject("WScript.Shell")
 shell.Run command,0
