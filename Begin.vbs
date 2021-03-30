Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\pi\Citi Bank program\Dependencies"
Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c start.bat"
oShell.Run strArgs, 0, false

