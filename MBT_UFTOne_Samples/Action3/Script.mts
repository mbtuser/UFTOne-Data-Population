﻿'wait(3)
'If Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Open").Exist(5) Then
'	Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Open").Click @@ script infofile_;_ZIP::ssf11.xml_;_
'End If @@ script infofile_;_ZIP::ssf8.xml_;_
Dim iURL
Dim objShell
iURL = "https://advantageonlinebanking.com/dashboard"
set objShell = CreateObject("Shell.Application")

Set fileSystemObj = createobject("Scripting.FileSystemObject")
edgeExist = "C:\Program Files\Mozilla Firefox\firefox.exe"
If fileSystemObj.FileExists(edgeExist) then
objShell.ShellExecute "C:\Program Files\Mozilla Firefox\firefox.exe", iURL, "", ""
Else
objShell.ShellExecute "C:\Program Files (x86)\Mozilla Firefox\firefox.exe", iURL, "", ""
End If
wait(3)
If Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link("Accounts").Exist(5) Then
wait(3)
	Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link("Accounts").Click
wait(3)
	Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click
	Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")
       Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
End If
wait(3)
