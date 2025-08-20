' =============================================================================
' ✅ פונקציה להצגת שכבת שגיאה ויזואלית על גבי דף האינטרנט.
' =============================================================================
Sub ShowErrorOnPage(errorMessage)
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
    On Error Resume Next
    Browser("micclass:=Browser").Page("micclass:=Page").RunScript(jsCode)
    On Error GoTo 0
    Wait(5) ' המתנה כדי שההודעה תוקלט בוידאו
End Sub

' --- START OF ORIGINAL ACTION LOGIC ---

Function GetObjectByNameSafe(logicalName)
    On Error Resume Next
    Set GetObjectByNameSafe = Nothing
    Select Case logicalName
        Case "username"
            Set GetObjectByNameSafe = Browser("CreationTime:=0").Page("title:=.*").WebEdit("name:=username")
        Case "password"
            Set GetObjectByNameSafe = Browser("CreationTime:=0").Page("title:=.*").WebEdit("name:=password")
        Case "signIn"
            Set GetObjectByNameSafe = Browser("CreationTime:=0").Page("title:=.*").WebButton("name:=Sign-In")
        Case "login"
            Set GetObjectByNameSafe = Browser("CreationTime:=0").Page("title:=.*").WebButton("name:=Login")
        Case "dashboardBtn"
            Set GetObjectByNameSafe = Browser("CreationTime:=0").Page("title:=.*").WebElement("innertext:=Bank Accounts")
    End Select
    On Error GoTo 0
End Function

Dim usernameObj, passwordObj, signInObj, loginObj, dashObj

Set usernameObj = GetObjectByNameSafe("username")
If Not usernameObj Is Nothing And usernameObj.Exist(3) Then
    usernameObj.Set "samples_demo_tests1"
Else
    ShowErrorOnPage "Username field not found"
    Reporter.ReportEvent micFail, "Element Validation", "Username field not found"
    ExitTest
End If

Set passwordObj = GetObjectByNameSafe("password")
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure "Aa1234567890"
Else
    ShowErrorOnPage "Password field not found"
    Reporter.ReportEvent micFail, "Element Validation", "Password field not found"
    ExitTest
End If

Set signInObj = GetObjectByNameSafe("signIn")
If signInObj Is Nothing Or Not signInObj.Exist(3) Then
    Set loginObj = GetObjectByNameSafe("login")
    If loginObj Is Nothing Or Not loginObj.Exist(3) Then
        ShowErrorOnPage "No login buttons (Sign-In or Login) were found"
        Reporter.ReportEvent micFail, "Element Validation", "No login buttons were found"
        ExitTest
    Else
        loginObj.Click
    End If
Else
    signInObj.Click
End If

Wait(3)

Set dashObj = GetObjectByNameSafe("dashboardBtn")
If Not dashObj Is Nothing And dashObj.Exist(10) Then
    dashObj.Click
Else
    ShowErrorOnPage "Dashboard button not found after login"
    Reporter.ReportEvent micFail, "Post-Login Validation", "Dashboard button not found"
    ExitTest
End If
