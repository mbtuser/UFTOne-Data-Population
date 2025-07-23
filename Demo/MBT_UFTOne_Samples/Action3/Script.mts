Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ✅ הודעת שגיאה HTA תמיד בראש המסך (Always on Top) ונשמרת 5 שניות
Sub ShowBlockingPopup(msg)
    On Error Resume Next
    Dim tempFilePath, f, safeMsg
    safeMsg = Replace(msg, """", "'")
    tempFilePath = "C:\Windows\Temp\popup_msg.hta"

    Set f = fileSystemObj.CreateTextFile(tempFilePath, True)
    f.WriteLine "<html><head><title>Error</title>"
    f.WriteLine "<hta:application showInTaskbar='yes' windowState='normal' sysMenu='no' caption='yes' border='thin' maximizeButton='no' minimizeButton='no' />"
    f.WriteLine "<script>"
    f.WriteLine "function setOnTop() {"
    f.WriteLine "  try {"
    f.WriteLine "    var shell = new ActiveXObject('WScript.Shell');"
    f.WriteLine "    shell.AppActivate(document.title);"
    f.WriteLine "  } catch(e) {}"
    f.WriteLine "  setTimeout('window.close()', 5000);"
    f.WriteLine "}"
    f.WriteLine "</script></head>"
    f.WriteLine "<body onload='setOnTop()' bgcolor='#fff0f0'>"
    f.WriteLine "<h2 style='color:red; font-family:sans-serif; text-align:center; margin-top:40px'>" & safeMsg & "</h2>"
    f.WriteLine "</body></html>"
    f.Close

    Dim wsh
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "mshta.exe """ & tempFilePath & """", 1, False
    Wait(6)
    Set wsh = Nothing
    On Error GoTo 0
End Sub

' ⏳ המתן אם קיימת תיקיית Report
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

