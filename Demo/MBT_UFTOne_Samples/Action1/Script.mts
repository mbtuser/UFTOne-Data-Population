Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ✅ הודעת שגיאה – קופצת Always on Top ונשמרת 5 שניות
Sub ShowPopupMessage(msg)
    On Error Resume Next
    Dim tempFilePath, f, safeMsg
    safeMsg = Replace(msg, """", "'")
    tempFilePath = "C:\Windows\Temp\uft_popup.hta"

    Set f = fileSystemObj.CreateTextFile(tempFilePath, True)
    f.WriteLine "<html><head><title>Error</title>"
    f.WriteLine "<hta:application showInTaskbar='yes' windowState='normal' sysMenu='no' caption='yes' border='thin' maximizeButton='no' minimizeButton='no' />"
    f.WriteLine "<script>"
    f.WriteLine "function setOnTop() {"
    f.WriteLine "  try { var shell = new ActiveXObject('WScript.Shell'); shell.AppActivate(document.title); } catch(e) {}"
    f.WriteLine "  setTimeout('window.close()', 5000);"
    f.WriteLine "}"
    f.WriteLine "</script></head>"
    f.WriteLine "<body onload='setOnTop()' bgcolor='#fff0f0'>"
    f.WriteLine "<h2 style='color:red; font-family:sans-serif; text-align:center; margin-top:40px'>" & safeMsg & "</h2>"
    f.WriteLine "</body></html>"
    f.Close

    CreateObject("WScript.Shell").Run "mshta.exe """ & tempFilePath & """", 1, False
    Wait(6)
    On Error GoTo 0
End Sub

' ⏳ המתנה אם תיקיית Report קיימת
Dim basePath, folder
basePath = "C:\test\repository\copy\src"
If fileSystemObj.FolderExists(basePath) Then
    For Each folder In fileSystemObj.GetFolder(basePath).SubFolders
        If InStr(folder.Name, "repo-") > 0 Then
            If fileSystemObj.FolderExists(folder.Path & "\repository\___mbt\_1\MBT_UFTOne_Samples_00001\Report") Then
                Wait(5)
                Exit For
            End If
        End If
    Next
End If

' ������ פתיחת דפדפן זמין
If fileSystemObj.FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    browserName = "chrome.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") Then
    browserPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    browserName = "msedge.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files\Mozilla Firefox\firefox.exe") Then
    browserPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
    browserName = "firefox.exe"
Else
    ShowPopupMessage "❌ No supported browser found"
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

' ������ מיפוי אובייקטים לפי שם לוגי
Function GetObjectByNameSafe(logicalName)
    On Error Resume Next
    Set GetObjectByNameSafe = Nothing
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

' ������ הכנסת שם משתמש
Dim userFieldName
userFieldName = Trim(Parameter("usernameField"))
If userFieldName = "" Then userFieldName = "username"

Set usernameObj = GetObjectByNameSafe(userFieldName)
If Not usernameObj Is Nothing And usernameObj.Exist(3) Then
    usernameObj.Set Parameter("username")
    Reporter.ReportEvent micPass, "Username", "Username set"
Else
    ShowPopupMessage "❌ Username field '" & userFieldName & "' not found"
    Reporter.ReportEvent micFail, "Username", "Username field missing"
    ExitTest
End If

' ������ הכנסת סיסמה
Dim passFieldName
passFieldName = Trim(Parameter("passwordField"))
If passFieldName = "" Then passFieldName = "password"

Set passwordObj = GetObjectByNameSafe(passFieldName)
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure Parameter("password")
    Reporter.ReportEvent micPass, "Password", "Password set"
Else
    ShowPopupMessage "❌ Password field '" & passFieldName & "' not found"
    Reporter.ReportEvent micFail, "Password", "Password field missing"
    ExitTest
End If

' ������ התחברות
Dim signInName, loginName
signInName = Trim(Parameter("signInButton"))
If signInName = "" Then signInName = "signIn"

Set signInObj = GetObjectByNameSafe(signInName)
If signInObj Is Nothing Or Not signInObj.Exist(3) Then
    loginName = Trim(Parameter("loginButton"))
    If loginName = "" Then loginName = "login"
    Set loginObj = GetObjectByNameSafe(loginName)
    If loginObj Is Nothing Or Not loginObj.Exist(3) Then
        ShowPopupMessage "❌ No login buttons found ('" & signInName & "' or '" & loginName & "')"
        Reporter.ReportEvent micFail, "Login", "No login buttons found"
        ExitTest
    Else
        loginObj.Click
    End If
Else
    signInObj.Click
End If

Wait(3)

' ✅ בדיקת דשבורד
Dim dashboardBtnName
dashboardBtnName = Trim(Parameter("dashboardButton"))
If dashboardBtnName = "" Then dashboardBtnName = "dashboardBtn"

Set dashObj = GetObjectByNameSafe(dashboardBtnName)
If Not dashObj Is Nothing And dashObj.Exist(20) Then
    Reporter.ReportEvent micPass, "Login Success", "Dashboard loaded"
    dashObj.Click
Else
    ShowPopupMessage "❌ Dashboard button '" & dashboardBtnName & "' not found"
    Reporter.ReportEvent micFail, "Login Failed", "Dashboard button missing"
    ExitTest
End If

Wait(3)
SystemUtil.CloseProcessByName browserName

