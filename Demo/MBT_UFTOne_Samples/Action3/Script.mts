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

            Reporter.ReportEvent micFail, "Account Creation", "Error: 'Name' input field for account creation not found. Displaying message on screen."
            ' Display on-screen error message for 5 seconds using DeviceReplay
            DeviceReplay.Screen.DrawText "ERROR: Account Name field not found!", 100, 100, 5, "red", 20
            Wait(5) 
        End If

    Else

        Reporter.ReportEvent micFail, "Account Creation", "Error: 'Open new account' button not found. Displaying message on screen."
        ' Display on-screen error message for 5 seconds using DeviceReplay
        DeviceReplay.Screen.DrawText "ERROR: 'Open new account' button not found!", 100, 150, 5, "red", 20
        Wait(5) 
    End If
Else

    Reporter.ReportEvent micFail, "Navigation", "Error: '" & accountsLinkText & "' link not found on dashboard. Displaying message on screen."
    ' Display on-screen error message for 5 seconds using DeviceReplay
    DeviceReplay.Screen.DrawText "ERROR: '" & accountsLinkText & "' link not found!", 100, 200, 5, "red", 20
    Wait(5) 
End If

Wait(3)


SystemUtil.CloseProcessByName browserName
