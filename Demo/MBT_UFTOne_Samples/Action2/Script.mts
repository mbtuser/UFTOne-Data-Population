Dim iURL, objShell, fileSystemObj, browserPath, browserName

iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ������ בדיקת דפדפן מותקן
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
    Reporter.ReportEvent micFail, "Browser Launch", "❌ No supported browser found"
    ExitTest
End If

' ������ פתיחת הדפדפן
objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

' ������ בדיקת אם כבר מחובר
If Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Exist(5) Then
    Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Click
    Wait(1)

    If Browser("Home - Advantage Bank_3").Page("Dashboard - Advantage").WebMenu("My Profile Management").Exist(3) Then
        Browser("Home - Advantage Bank_3").Page("Dashboard - Advantage").WebMenu("My Profile Management").Select "Logout"
        Reporter.ReportEvent micPass, "Logout", "✅ User logged out successfully"
    Else
        Reporter.ReportEvent micWarning, "Logout Menu", "⚠️ Logout menu not found"
    End If
Else
    Reporter.ReportEvent micDone, "Login State", "ℹ️ User not logged in – proceeding to registration"
End If

Wait(2)

' ������ רישום משתמש חדש
If Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Registration").Exist(5) Then
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Registration").Click
    Wait(3)

    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("firstName").Set Parameter("firstName")
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("lastName").Set Parameter("lastName")
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("email").Set Parameter("email")
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("username").Set Parameter("username")
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("password").Set Parameter("password")
    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("confirmPass").Set Parameter("confirmPass")

    Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Register").Click
    Reporter.ReportEvent micDone, "Registration", "������ Registration submitted"
Else
    Reporter.ReportEvent micFail, "Registration", "❌ Registration button not found"
    ExitTest
End If

Wait(5)

' ������ סגירת הדפדפן לאחר השהייה למניעת שגיאות IO
Wait(3)
SystemUtil.CloseProcessByName browserName

