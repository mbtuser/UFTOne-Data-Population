' ✅ הודעת שגיאה שתוקלט
Sub ShowPopupMessage(msg)
    On Error Resume Next
    msg = Replace(msg, """", "'")
    MsgBox msg, 48, "❌ UFT Error"
    On Error GoTo 0
End Sub

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
    ShowPopupMessage "❌ Username field not found"
    ExitTest
End If

Set passwordObj = GetObjectByNameSafe("password")
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure "Aa1234567890"
Else
    ShowPopupMessage "❌ Password field not found"
    ExitTest
End If

Set signInObj = GetObjectByNameSafe("signIn")
If signInObj Is Nothing Or Not signInObj.Exist(3) Then
    Set loginObj = GetObjectByNameSafe("login")
    If loginObj Is Nothing Or Not loginObj.Exist(3) Then
        ShowPopupMessage "❌ No login buttons found"
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
    ShowPopupMessage "❌ Dashboard button not found"
    ExitTest
End If

