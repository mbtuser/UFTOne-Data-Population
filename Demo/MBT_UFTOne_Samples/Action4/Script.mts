' === Logout ===
Option Explicit

Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' בדיקת דפדפן זמין
If fileSystemObj.FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    browserName = "chrome.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") Then
    browserPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    browserName = "msedge.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files\Mozilla Firefox\firefox.exe") Then
    browserPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
    browserName = "firefox.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") Then
    browserPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    browserName = "firefox.exe"
Else
    OverlayFail "Browser Launch", "No supported browser found on this machine"
End If

' ⭐ כאן היה Typo ("Lo) – תוקן:
objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(3)

' בדיקת האם המשתמש מחובר (כפתור תפריט קיים)
If Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Exist(5) Then
    Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Click
    Wait(1)

    If Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebMenu("My Profile Management").Exist(3) Then
        Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebMenu("My Profile Management").Select "Logout"
        Reporter.ReportEvent micPass, "Logout", "User logged out successfully"
    Else
        OverlayFail "Logout", "Logout menu not found"
    End If
Else
    Reporter.ReportEvent micDone, "Login Status", "User is not logged in – Dashboard button not found"
End If

Wait(3)
SystemUtil.CloseProcessByName browserName

