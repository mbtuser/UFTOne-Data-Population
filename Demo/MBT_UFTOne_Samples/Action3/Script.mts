Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

' ✅ הודעת שגיאה שתוקלט ב־UFT עם MsgBox חוסם
Sub ShowPopupMessage(msg)
    On Error Resume Next
    msg = Replace(msg, """", "'")
    MsgBox msg, 48, "❌ UFT Error"
    On Error GoTo 0
End Sub

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
    ShowPopupMessage "❌ No supported browser found"
    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"
    ExitTest
End If

objShell.ShellExecute browserPath, iURL, "", "", 1
Wait(6)

' ������ ניסיון ללחוץ על לינק Accounts
Dim accountsLinkText, linkDesc, linkCount, matchingLinks
On Error Resume Next
accountsLinkText = Trim(Parameter("ElementName"))
If accountsLinkText = "" Then accountsLinkText = "Accounts"
On Error GoTo 0

Set linkDesc = Description.Create()
linkDesc("micclass").Value = "Link"
linkDesc("innertext").Value = accountsLinkText

Set matchingLinks = Browser("CreationTime:=0").Page("title:=.*").ChildObjects(linkDesc)
If Not matchingLinks Is Nothing Then
    linkCount = matchingLinks.Count
Else
    linkCount = 0
End If

If linkCount > 0 Then
    matchingLinks(0).Click
    Wait(2)
    Reporter.ReportEvent micPass, "Navigation", "Clicked '" & accountsLinkText & "'"
Else
    ShowPopupMessage "❌ Link '" & accountsLinkText & "' not found"
    Reporter.ReportEvent micFail, "Navigation", "Link '" & accountsLinkText & "' not found"
    ExitTest
End If

SystemUtil.CloseProcessByName browserName

