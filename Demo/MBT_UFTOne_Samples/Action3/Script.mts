Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ������ פונקציית הודעת שגיאה נראית ומוקלטת בוודאות
Sub ShowBlockingPopup(msg)
    Dim cmd, tempVbsFile
    tempVbsFile = "C:\Windows\Temp\blocking_popup.vbs"
    
    msg = Replace(msg, """", "'")
    
    ' צור קובץ .vbs שמציג הודעה חוסמת (blocking) שנשארת על המסך
    With fileSystemObj.CreateTextFile(tempVbsFile, True)
        .WriteLine "MsgBox """ & msg & """, 48, ""❌ UFT Automation Error"""
        .Close
    End With

    ' הפעל את ההודעה בצורה חוסמת כדי לוודא שהיא תישאר פתוחה לרגע ותוקלט
    Dim wsh
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "wscript.exe """ & tempVbsFile & """", 1, True
    Set wsh = Nothing
End Sub

' ⏳ המתנה אם קיימת תיקיית ריצה פעילה
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
    ShowBlockingPopup "❌ No supported browser found on this machine."
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(6)

' ������ ניתוח טקסט הלינק
Dim accountsLinkText, linkDesc, linkCount
accountsLinkText = Trim(Parameter("ElementName"))
If accountsLinkText = "" Then accountsLinkText = "Accounts"

Set linkDesc = Description.Create()
linkDesc("micclass").Value = "Link"
linkDesc("innertext").Value = accountsLinkText

Dim accountsPage, matchingLinks
Set accountsPage = Browser("Dashboard - Advantage").Page("Dashboard - Advantage")

On Error Resume Next
Set matchingLinks = accountsPage.ChildObjects(linkDesc)
linkCount = matchingLinks.Count
On Error GoTo 0

If linkCount > 0 Then
    matchingLinks(0).Click
    Wait(3)

    ' ������ פתיחת חשבון חדש
    If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Exist(5) Then
        Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Open new account").Click

        If Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Exist(3) Then
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebEdit("name").Set Parameter("accountName")
            Browser("Dashboard - Advantage").Page("Accounts - Advantage Bank").WebButton("Create").Click
            Reporter.ReportEvent micPass, "Account Creation", "New account created successfully"
        Else
            ShowBlockingPopup "❌ Input field for 'name' was not found."
            Reporter.ReportEvent micFail, "Account Creation", "'name' input field not found"
            ExitTest
        End If
    Else
        ShowBlockingPopup "❌ 'Open new account' button not found."
        Reporter.ReportEvent micFail, "Account Creation", "'Open new account' button missing"
        ExitTest
    End If
Else
    ShowBlockingPopup "❌ Link '" & accountsLinkText & "' not found on dashboard."
    Reporter.ReportEvent micFail, "Navigation", "Link '" & accountsLinkText & "' not found"
    ExitTest
End If

Wait(4)
SystemUtil.CloseProcessByName browserName

