' =============================================================================
' ✅ פונקציה להצגת שכבת שגיאה ויזואלית על גבי דף האינטרנט.
' =============================================================================
Sub ShowErrorOnPage(errorMessage)
    On Error Resume Next
    Dim jsErrorMessage
    jsErrorMessage = Replace(errorMessage, "'", "\'")
    jsErrorMessage = Replace(jsErrorMessage, """", "\""")
    jsErrorMessage = Replace(jsErrorMessage, vbCrLf, "<br>")
    Dim jsCode
    jsCode = "var overlayDiv = document.createElement('div');" & _
             "overlayDiv.id = 'uft-error-overlay';" & _
             "overlayDiv.style.cssText = 'position:fixed; top:30px; left:50%; transform:translateX(-50%); padding:25px; background-color:rgba(220, 53, 69, 0.9); color:white; font-size:22px; font-weight:bold; border:4px solid black; border-radius:12px; z-index:999999; box-shadow: 0 0 20px rgba(0,0,0,0.7); font-family:Arial,sans-serif; text-align:center;';" & _
             "overlayDiv.innerHTML = '❌ UFT One Test Failure ❌<hr style=""border-color:white; margin:10px 0;""><p style=""font-size:18px; font-weight:normal;"">" & jsErrorMessage & "</p>';" & _
             "if (document.body) { document.body.appendChild(overlayDiv); } else { console.error('UFT Error: " & jsErrorMessage & "'); }"
    Browser("micclass:=Browser").Page("micclass:=Page").RunScript(jsCode)
    Wait(5) ' המתנה כדי שההודעה תוקלט בוידאו
    On Error GoTo 0
End Sub

' --- START OF ORIGINAL ACTION LOGIC ---

Dim iURL, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

If fileSystemObj.FileExists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    browserName = "chrome.exe"
ElseIf fileSystemObj.FileExists("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") Then
    browserPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    browserName = "msedge.exe"
Else
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

' שימוש בפקודה אמינה יותר של UFT לפתיחת דפדפן
SystemUtil.Run browserPath, iURL
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

Dim usernameObj, passwordObj, signInObj, loginObj, dashObj

Set usernameObj = GetObjectByName(Parameter("usernameField"))
If Not usernameObj Is Nothing And usernameObj.Exist(3) Then
    usernameObj.Set Parameter("username")
    Reporter.ReportEvent micPass, "Username Set", "Username set successfully"
Else
    ShowErrorOnPage "Failed to find username field using identifier: '" & Parameter("usernameField") & "'"
    Reporter.ReportEvent micFail, "Username Not Found", "Failed to find username field"
    ExitTest
End If

Set passwordObj = GetObjectByName(Parameter("passwordField"))
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure Parameter("password")
    Reporter.ReportEvent micPass, "Password Set", "Password set successfully"
Else
    ShowErrorOnPage "Failed to find password field using identifier: '" & Parameter("passwordField") & "'"
    Reporter.ReportEvent micFail, "Password Not Found", "Failed to find password field"
    ExitTest
End If

Set signInObj = GetObjectByName(Parameter("signInButton"))
Set loginObj  = GetObjectByName(Parameter("loginButton"))

If Not signInObj Is Nothing And signInObj.Exist(3) Then
    signInObj.Click
ElseIf Not loginObj Is Nothing And loginObj.Exist(3) Then
    loginObj.Click
Else
    ShowErrorOnPage "No login button found using identifiers: '" & Parameter("signInButton") & "' or '" & Parameter("loginButton") & "'"
    Reporter.ReportEvent micFail, "Login Button", "No login button found"
    ExitTest
End If

Wait(3)

Set dashObj = GetObjectByName(Parameter("dashboardButton"))
If Not dashObj Is Nothing And dashObj.Exist(20) Then
    Reporter.ReportEvent micPass, "Login Test", "Login successful"
    dashObj.Click
Else
    ShowErrorOnPage "Login failed. Dashboard button not found using identifier: '" & Parameter("dashboardButton") & "'"
    Reporter.ReportEvent micFail, "Login Test", "Login failed"
    ExitTest
End If

SystemUtil.CloseProcessByName browserName
