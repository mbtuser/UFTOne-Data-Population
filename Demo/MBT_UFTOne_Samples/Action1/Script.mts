' אותו ShowPopupMessage כמו למעלה
Sub ShowPopupMessage(msg)
    On Error Resume Next
    Dim tempVbsPath, f
    tempVbsPath = "C:\Windows\Temp\uft_error_popup.vbs"
    msg = Replace(msg, """", "'")
    Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(tempVbsPath, True)
    f.WriteLine "MsgBox """ & msg & """, 48, ""❌ UFT Error"""
    f.Close
    CreateObject("WScript.Shell").Run "wscript.exe """ & tempVbsPath & """", 1, True
    On Error GoTo 0
End Sub

Dim iURL, objShell, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")

If CreateObject("Scripting.FileSystemObject").FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    browserName = "chrome.exe"
ElseIf CreateObject("Scripting.FileSystemObject").FileExists("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") Then
    browserPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    browserName = "msedge.exe"
ElseIf CreateObject("Scripting.FileSystemObject").FileExists("C:\Program Files\Mozilla Firefox\firefox.exe") Then
    browserPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
    browserName = "firefox.exe"
Else
    ShowPopupMessage "❌ No supported browser found"
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(6)

Function GetObjectByNameSafe(logicalName)
    On Error Resume Next
    Select Case logicalName
        Case "username"
            Set GetObjectByNameSafe = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("username")
        Case "password"
            Set GetObjectByNameSafe = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("password")
        Case "signIn"
            Set GetObjectByNameSafe = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Sign-In")
        Case "login"
            Set GetObjectByNameSafe = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Login")
        Case "dashboardBtn"
            Set GetObjectByNameSafe = Browser("Dashboard - Advantage_2").Page("Dashboard - Advantage").WebElement("Bank Accounts")
    End Select
    On Error GoTo 0
End Function

Dim usernameObj, passwordObj
Set usernameObj = GetObjectByNameSafe("username")
If usernameObj Is Nothing Or Not usernameObj.Exist(3) Then
    ShowPopupMessage "❌ Username field not found"
    ExitTest
End If
usernameObj.Set Parameter("username")

Set passwordObj = GetObjectByNameSafe("password")
If passwordObj Is Nothing Or Not passwordObj.Exist(3) Then
    ShowPopupMessage "❌ Password field not found"
    ExitTest
End If
passwordObj.SetSecure Parameter("password")

Dim loginObj
Set loginObj = GetObjectByNameSafe("login")
If loginObj Is Nothing Or Not loginObj.Exist(3) Then
    ShowPopupMessage "❌ Login button not found"
    ExitTest
End If
loginObj.Click
Wait(4)

Dim dashObj
Set dashObj = GetObjectByNameSafe("dashboardBtn")
If dashObj Is Nothing Or Not dashObj.Exist(10) Then
    ShowPopupMessage "❌ Dashboard element not found"
    ExitTest
Else
    dashObj.Click
End If

Wait(4)
SystemUtil.CloseProcessByName browserName

