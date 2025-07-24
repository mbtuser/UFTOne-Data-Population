Dim iURL, objShell, fileSystemObj, browserPath, browserName


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

    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found on this machine"

    ExitTest
End If


objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

' Function to inject an error message into the web page
Function InjectWebErrorMessage(msgText)
    Dim jsCode
    jsCode = "var errorDiv = document.createElement('div');" & _
             "errorDiv.id = 'ciErrorMessage';" & _
             "errorDiv.style.position = 'fixed';" & _
             "errorDiv.style.top = '10px';" & _
             "errorDiv.style.right = '10px';" & _
             "errorDiv.style.backgroundColor = 'red';" & _
             "errorDiv.style.color = 'white';" & _
             "errorDiv.style.padding = '15px';" & _
             "errorDiv.style.border = '2px solid darkred';" & _
             "errorDiv.style.borderRadius = '8px';" & _
             "errorDiv.style.zIndex = '99999';" & _
             "errorDiv.style.fontSize = '18px';" & _
             "errorDiv.style.fontWeight = 'bold';" & _
             "errorDiv.style.animation = 'fadeIn 0.5s';" & _
             "errorDiv.innerHTML = '" & Replace(msgText, "'", "\'") & "';" & _
             "var existingError = document.getElementById('ciErrorMessage');" & _
             "if (existingError) { existingError.parentNode.removeChild(existingError); }" & _
             "document.body.appendChild(errorDiv);" & _
             "setTimeout(function() { if(errorDiv.parentNode) errorDiv.parentNode.removeChild(errorDiv); }, 5000);" ' Remove after 5 seconds

    On Error Resume Next ' In case the browser/page is not ready or valid
    Browser("micClass:=Browser").Page("micClass:=Page").RunScript jsCode
    On Error GoTo 0
End Function


Dim accountsLinkText

accountsLinkText = Parameter("ElementName")

If Trim(accountsLinkText) = "" Then

    accountsLinkText = "Accounts"
End If

If Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link(accountsLinkText).Exist(5) Then

    Wait(3)

    Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link("Accounts").Click

    Wait(3)


    If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Exist(3) Then

        Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click


        If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Exist(3) Then

            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")

            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click

            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"

        Else

            Reporter.ReportEvent micFail, "Account Creation", "ERROR: 'Name' input field for account creation not found. Displaying message on web page."
            InjectWebErrorMessage "ERROR: Account Name field not found!"
            Wait(5) 
        End If

    Else

        Reporter.ReportEvent micFail, "Account Creation", "ERROR: 'Open new account' button not found. Displaying message on web page."
        InjectWebErrorMessage "ERROR: 'Open new account' button not found!"
        Wait(5) 
    End If
Else

    Reporter.ReportEvent micFail, "Navigation", "ERROR: '" & accountsLinkText & "' link not found on dashboard. Displaying message on web page."
    InjectWebErrorMessage "ERROR: '" & accountsLinkText & "' link not found!"
    Wait(5) 
End If

Wait(3)


SystemUtil.CloseProcessByName browserName
