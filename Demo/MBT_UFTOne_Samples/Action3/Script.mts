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

' Function to inject a visual error message into the current web page
Function InjectWebErrorMessage(msgText)
    Dim jsCode
    ' JavaScript code to create a styled DIV element, add it to the body, and remove it after 5 seconds
    jsCode = "var errorDiv = document.createElement('div');" & _
             "errorDiv.id = 'ciErrorOverlay';" & _
             "errorDiv.style.position = 'fixed';" & _
             "errorDiv.style.top = '10%'; " & _
             "errorDiv.style.left = '50%'; " & _
             "errorDiv.style.transform = 'translate(-50%, -50%)';" & _
             "errorDiv.style.backgroundColor = 'red';" & _
             "errorDiv.style.color = 'white';" & _
             "errorDiv.style.padding = '20px 30px';" & _
             "errorDiv.style.border = '3px solid darkred';" & _
             "errorDiv.style.borderRadius = '10px';" & _
             "errorDiv.style.zIndex = '99999';" & _
             "errorDiv.style.fontSize = '24px';" & _
             "errorDiv.style.fontWeight = 'bold';" & _
             "errorDiv.style.textAlign = 'center';" & _
             "errorDiv.style.boxShadow = '0 0 15px rgba(0,0,0,0.5)';" & _
             "errorDiv.style.opacity = '0';" & _
             "errorDiv.style.transition = 'opacity 0.5s ease-in-out';" & _
             "errorDiv.innerHTML = '" & Replace(msgText, "'", "\'") & "';" & _
             "var existingError = document.getElementById('ciErrorOverlay');" & _
             "if (existingError) { existingError.parentNode.removeChild(existingError); }" & _
             "document.body.appendChild(errorDiv);" & _
             "setTimeout(function() { errorDiv.style.opacity = '1'; }, 100);" & _
             "setTimeout(function() { " & _
             "    if(errorDiv.parentNode) { " & _
             "        errorDiv.style.opacity = '0';" & _
             "        setTimeout(function() { if(errorDiv.parentNode) errorDiv.parentNode.removeChild(errorDiv); }, 500); " & _
             "    }" & _
             "}, 5000);" ' Show for 5 seconds

    On Error Resume Next ' Use On Error Resume Next for robustness if the page is not fully loaded
    ' Target the current active browser and page
    Browser("micClass:=Browser").Page("micClass:=Page").RunScript jsCode
    If Err.Number <> 0 Then
        Reporter.ReportEvent micWarning, "JavaScript Injection", "Could not inject error message: " & Err.Description
    End If
    On Error GoTo 0
End Function


Dim accountsLinkText
accountsLinkText = Parameter("ElementName")

If Trim(accountsLinkText) = "" Then
    accountsLinkText = "Accounts"
End If

' *** שינוי חשוב כאן: הגדרת האלמנט באופן פרוגרמטי ***
' במקום Link(accountsLinkText) שמחפש ברפוזיטורי, נגדיר תכונות כמו innertext
Set objAccountsLink = Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link("innertext:=" & accountsLinkText, "micClass:=Link")

If objAccountsLink.Exist(5) Then ' עכשיו בדיקת Exist תתבצע על אלמנט שנמצא על בסיס תכונות
    Wait(3)
    objAccountsLink.Click ' נשתמש באובייקט שזיהינו
    Wait(3)

    ' עבור כפתור "Open new account" - נגדיר אותו גם פרוגרמטית (אם לא מופיע ברפוזיטורי)
    Set objOpenNewAccountButton = Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("innertext:=Open new account", "micClass:=WebButton")
    If objOpenNewAccountButton.Exist(3) Then
        objOpenNewAccountButton.Click

        ' עבור שדה "name" - נגדיר אותו גם פרוגרמטית
        Set objAccountNameField = Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name:=name", "micClass:=WebEdit")
        If objAccountNameField.Exist(3) Then
            objAccountNameField.Set Parameter("accountName")
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"
        Else
            Reporter.ReportEvent micFail, "Account Creation", "ERROR: 'Name' input field for account creation not found. Injected message to web page."
            InjectWebErrorMessage "ERROR: Account Name field not found!"
            Wait(5) ' Keep the wait to ensure message is visible in recording
        End If
    Else
        Reporter.ReportEvent micFail, "Account Creation", "ERROR: 'Open new account' button not found. Injected message to web page."
        InjectWebErrorMessage "ERROR: 'Open new account' button not found!"
        Wait(5) ' Keep the wait to ensure message is visible in recording
    End If
Else
    Reporter.ReportEvent micFail, "Navigation", "ERROR: '" & accountsLinkText & "' link not found on dashboard. Injected message to web page."
    InjectWebErrorMessage "ERROR: '" & accountsLinkText & "' link not found!"
    Wait(5) ' Keep the wait to ensure message is visible in recording
End If

Wait(3)


SystemUtil.CloseProcessByName browserName
