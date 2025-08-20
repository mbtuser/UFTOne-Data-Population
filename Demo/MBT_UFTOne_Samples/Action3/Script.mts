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
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found on this machine"
    ExitTest
End If

SystemUtil.Run browserPath, iURL
Wait(5)

Dim accountsLinkText
accountsLinkText = Parameter("ElementName")
If Trim(accountsLinkText) = "" Then
    accountsLinkText = "Accounts"
End If

If Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link(accountsLinkText).Exist(5) Then
    Wait(3)
    Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link(accountsLinkText).Click
    Wait(3)

    If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Exist(3) Then
        Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click

        If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Exist(3) Then
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"
        Else
            ShowErrorOnPage "Name input field not found on 'New Account' page."
            Reporter.ReportEvent micFail, "Account Creation", "Name input field not found"
            ExitTest
        End If
    Else
        ShowErrorOnPage "'Open new account' button not found on 'Accounts' page."
        Reporter.ReportEvent micFail, "Account Creation", "'Open new account' button not found"
        ExitTest
    End If
Else
    ShowErrorOnPage "'Accounts' link not found on dashboard using text: '" & accountsLinkText & "'"
    Reporter.ReportEvent micFail, "Navigation", "'Accounts' link not found on dashboard"
    ExitTest
End If

Wait(3)
SystemUtil.CloseProcessByName browserName
