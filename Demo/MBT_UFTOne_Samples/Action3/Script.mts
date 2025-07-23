Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ������ פונקציה להצגת הודעה על המסך (לא חוסמת, נראית בהקלטה)
Function ShowPopupMessage(msg)
    On Error Resume Next
    Dim tempFilePath, f, safeMsg
    safeMsg = Replace(msg, """", "'")
    tempFilePath = "C:\Windows\Temp\msg_popup.vbs"
    
    Set f = fileSystemObj.CreateTextFile(tempFilePath, True)
    If Not f Is Nothing Then
        f.WriteLine "Set WshShell = CreateObject(""WScript.Shell"")"
        f.WriteLine "WshShell.Popup """ & safeMsg & """, 7, ""❌ Error"", 48"
        f.Close
        CreateObject("WScript.Shell").Run "wscript.exe """ & tempFilePath & """", 1, False
    Else
        Reporter.ReportEvent micWarning, "Popup Failure", "⚠ Could not create popup file"
    End If
    On Error GoTo 0
End Function

' ������ בדיקה אם תיקיית Report קיימת (מסמלת ריצה קודמת פעילה)
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
ElseIf fileSystemObj.FileExists("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") Then
    browserPath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    browserName = "firefox.exe"
Else
    ShowPopupMessage "❌ No supported browser found on this machine"
    Reporter.ReportEvent micFail, "Browser Launch", "No browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(5)

' ������️ בדיקה אם קיים הלינק לפי הטקסט שהוזן בפרמטר
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
            ShowPopupMessage "❌ Input field for 'name' was not found"
            Reporter.ReportEvent micFail, "Account Creation", "Missing 'name' input field"
            ExitTest
        End If
    Else
        ShowPopupMessage "❌ 'Open new account' button not found on the page"
        Reporter.ReportEvent micFail, "Account Creation", "'Open new account' button missing"
        ExitTest
    End If
Else
    ShowPopupMessage "❌ Link '" & accountsLinkText & "' not found on dashboard"
    Reporter.ReportEvent micFail, "Navigation", "'" & accountsLinkText & "' link not found"
    ExitTest
End If

Wait(4)
SystemUtil.CloseProcessByName browserName

