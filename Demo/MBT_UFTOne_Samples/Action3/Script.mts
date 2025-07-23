Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ������ פונקציה להצגת הודעת שגיאה זמנית (לא חוסמת)
Function ShowPopupMessage(msg)
    On Error Resume Next
    Dim tempFilePath, f, safeMsg
    safeMsg = Replace(msg, """", "'") ' מניעת גרשיים כפולים שגורמים לקריסה
    tempFilePath = "C:\Windows\Temp\msg.vbs"

    Set f = fileSystemObj.CreateTextFile(tempFilePath, True)
    If Not f Is Nothing Then
        f.WriteLine "Set oShell = CreateObject(""WScript.Shell"")"
        f.WriteLine "oShell.Popup """ & safeMsg & """, 5, ""❌ Element Not Found"", 48"
        f.Close
        CreateObject("WScript.Shell").Run "wscript.exe """ & tempFilePath & """", 1, False
    Else
        Reporter.ReportEvent micWarning, "Popup Failure", "⚠ Could not create popup script file"
    End If
    On Error GoTo 0
End Function

' ������ המתנה אם קיימת תיקיית דוח נעולה כלשהי
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

' ������️ בדיקת שם הלינק מתוך פרמטר
Dim accountsLinkText
accountsLinkText = Parameter("ElementName")
If Trim(accountsLinkText) = "" Then accountsLinkText = "Accounts"

' ������ ניווט לחשבונות
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

' ⏳ השהייה לפני סגירת הדפדפן – שיאפשר הקלטת הודעה
Wait(3)
SystemUtil.CloseProcessByName browserName

