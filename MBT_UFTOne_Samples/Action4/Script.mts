﻿'If Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Open").Exist Then
'       Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Open").Click @@ script infofile_;_ZIP::ssf9.xml_;_
'End If
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
wait(3) @@ script infofile_;_ZIP::ssf13.xml_;_
       Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Click @@ script infofile_;_ZIP::ssf2.xml_;_
       Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebMenu("My Profile Management").Select "Logout" @@ script infofile_;_ZIP::ssf3.xml_;_
        systemUtil.CloseProcessByName ("firefox.exe")
