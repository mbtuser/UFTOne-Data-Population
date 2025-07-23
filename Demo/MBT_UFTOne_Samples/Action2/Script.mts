Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ✅ הודעת שגיאה – MsgBox מתוך קובץ VBS חוסם
Sub ShowPopupMessage(msg)
    On Error Resume Next
    Dim tempVbsPath, f
    tempVbsPath = "C:\Windows\Temp\uft_error_popup.vbs"
    msg = Replace(msg, """", "'") ' בטיחות למחרוזת

    Set f = fileSystemObj.CreateTextFile(tempVbsPath, True)
    f.WriteLine "MsgBox """ & msg & """, 48, ""❌ UFT Error"""
    f.Close

    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe """ & tempVbsPath & """", 1, True ' True = חוסם
    Set shell = Nothing
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
    ShowPopupMessage "❌ No supported browser found on this machine."
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(6)

' ������ ניסיון ללחוץ על לינק Accounts
Dim accountsLinkText, linkDesc, linkCount, matchingLinks
accountsLinkText = Trim(Parameter("ElementName"))
If accountsLinkText = "" Then accountsLinkText = "Accounts"

Set linkDesc = Description.Create()
linkDesc("micclass").Value = "Link"
linkDesc("innertext").Value = accountsLinkText

Set matchingLinks = Browser("Dashboard - Advantage").Page("Dashboard - Advantage").ChildObjects(linkDesc)
linkCount = matchingLinks.Count

If linkCount > 0 Then
    matchingLinks(0).Click
    Wait(3)

    ' ������ לחיצה על כפתור יצירת חשבון
    If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Exist(5) Then
        Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click

        If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Exist(3) Then
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"
        Else
            ShowPopupMessage "❌ Input field for 'name' was not found."
            Reporter.ReportEvent micFail, "Account Creation", "'name' input field not found"
            ExitTest
        End If
    Else
        ShowPopupMessage "❌ 'Open new account' button not found."
        Reporter.ReportEvent micFail, "Account Creation", "'Open new account' button missing"
        ExitTest
    End If
Else
    ShowPopupMessage "❌ Link '" & accountsLinkText & "' not found on dashboard."
    Reporter.ReportEvent micFail, "Navigation", "Link '" & accountsLinkText & "' not found"
    ExitTest
End If

Wait(4)
SystemUtil.CloseProcessByName browserName

