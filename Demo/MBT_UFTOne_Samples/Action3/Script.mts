Dim iURL, objShell, fileSystemObj, browserPath, browserName

iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ������ פונקציה להצגת הודעת שגיאה זמנית
Function ShowPopupMessage(msg)
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Popup msg, 5, "❌ Element Not Found", 48 ' 48 = אייקון אזהרה
End Function

' ������ אם תיקיית הדוח קיימת, המתן לפני התחלת הריצה
If fileSystemObj.FolderExists("C:\test\repository\copy\src\repo-1006\repository\___mbt\_1\MBT_UFTOne_Samples_00001\Report") Then
    Wait(5)
End If

' ������ פתיחת דפדפן לפי מה שמותקן
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

' ������ ניווט לדף החשבונות
If Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link(accountsLinkText).Exist(5) Then
    Wait(3)
    Browser("Dashboard - Advantage").Page("Dashboard - Advantage").Link(accountsLinkText).Click
    Wait(3)

    ' ������ פתיחת חשבון חדש
    If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Exist(3) Then
        Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click

        If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Exist(3) Then
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"
        Else
            ShowPopupMessage "❌ The element 'name' input field was not found on the page."
            Reporter.ReportEvent micFail, "Account Creation", "Name input field not found"
            ExitTest
        End If
    Else
        ShowPopupMessage "❌ The button 'Open new account' was not found on the page."
        Reporter.ReportEvent micFail, "Account Creation", "'Open new account' button not found"
        ExitTest
    End If
Else
    ShowPopupMessage "❌ The link '" & accountsLinkText & "' was not found on the dashboard page."
    Reporter.ReportEvent micFail, "Navigation", "'" & accountsLinkText & "' link not found on dashboard"
    ExitTest
End If

' ✅ המתן לפני סגירת הדפדפן למניעת בעיות כתיבת דוח
Wait(3)
SystemUtil.CloseProcessByName browserName

