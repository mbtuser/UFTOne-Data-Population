﻿Dim iURL, objShell, fileSystemObj, browserPath, browserName

iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

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
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

Function GetObjectByName(elementName)
    Select Case elementName
        Case "username"
            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("username")
        Case "password"
            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("password")
        Case "signIn"
            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Sign-In")
        Case "login"
            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Login")
        Case "dashboardBtn"
            Set GetObjectByName = Browser("Dashboard - Advantage_2").Page("Dashboard - Advantage").WebElement("Bank Accounts")
        Case Else
            Set GetObjectByName = Nothing
    End Select
End Function

Set usernameObj = GetObjectByName(Parameter("usernameField"))
If Not usernameObj Is Nothing And usernameObj.Exist(3) Then
    usernameObj.Set Parameter("username")
    Reporter.ReportEvent micPass, "Username Set", "Username set successfully"
Else
    MsgBox "❌ The element 'usernameField' was not found on the page.", vbCritical, "Element Not Found"
    Reporter.ReportEvent micFail, "Username Not Found", "Failed to find username field"
    ExitTest
End If

Set passwordObj = GetObjectByName(Parameter("passwordField"))
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure Parameter("password")
    Reporter.ReportEvent micPass, "Password Set", "Password set successfully"
Else
    MsgBox "❌ The element 'passwordField' was not found on the page.", vbCritical, "Element Not Found"
    Reporter.ReportEvent micFail, "Password Not Found", "Failed to find password field"
    ExitTest
End If

Set signInObj = GetObjectByName(Parameter("signInButton"))
If signInObj Is Nothing Or Not signInObj.Exist(3) Then
    Set loginObj = GetObjectByName(Parameter("loginButton"))
    If loginObj Is Nothing Or Not loginObj.Exist(3) Then
        MsgBox "❌ Neither 'signInButton' nor 'loginButton' was found on the page.", vbCritical, "Element Not Found"
        Reporter.ReportEvent micFail, "Login Button", "No login button found"
        ExitTest
    Else
        loginObj.Click
    End If
Else
    signInObj.Click
End If

Wait(3)

Set dashObj = GetObjectByName(Parameter("dashboardButton"))
If Not dashObj Is Nothing And dashObj.Exist(20) Then
    Reporter.ReportEvent micPass, "Login Test", "Login successful"
    dashObj.Click
Else
    MsgBox "❌ The element 'dashboardButton' was not found on the page.", vbCritical, "Element Not Found"
    Reporter.ReportEvent micFail, "Login Test", "Login failed"
    ExitTest
End If

'SystemUtil.CloseProcessByName browserName
