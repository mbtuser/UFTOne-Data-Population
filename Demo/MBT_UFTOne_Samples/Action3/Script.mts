' =============================================================================
' ✅ פונקציה להצגת שכבת שגיאה ויזואלית על גבי דף האינטרנט.
' =============================================================================
Sub ShowErrorOnPage(errorMessage)
    Dim jsErrorMessage
    jsErrorMessage = Replace(errorMessage, "'", "\'")
    jsErrorMessage = Replace(jsErrorMessage, """", "\""")
    jsErrorMessage = Replace(jsErrorMessage, vbCrLf, "<br>")
    Dim jsCode
    jsCode = "var overlayDiv = document.createElement('div');" & _
             "overlayDiv.id = 'uft-error-overlay';" & _
             "overlayDiv.style.cssText = 'position:fixed; top:30px; left:50%; transform:translateX(-50%); padding:25px; background-color:rgba(220, 53, 69, 0.9); color:white; font-size:22px; font-weight:bold; border:4px solid black; border-radius:12px; z-index:999999; box-shadow: 0 0 20px rgba(0,0,0,0.7); font-family:Arial,sans-serif; text-align:center;';" & _
             "overlayDiv.innerHTML = '❌ UFT One Test Failure ❌<hr style=""border-color:white; margin:10px 0;""><p style=""font-size:18px; font-weight:normal;"">" & jsErrorMessage & "</p>';" & _
             "if (document.body) { document.body.appendChild(overlayDiv); } else { console.error('UFT Error: " & jsErrorMessage & "'); }"
    On Error Resume Next
    Browser("micclass:=Browser").Page("micclass:=Page").RunScript(jsCode)
    On Error GoTo 0
    Wait(5) ' המתנה כדי שההודעה תוקלט בוידאו
End Sub

' --- START OF ORIGINAL ACTION LOGIC ---

Dim iURL, objShell, fileSystemObj, browserPath, browserName
iURL = "https://advantageonlinebanking.com/dashboard"
Set objShell = CreateObject("Shell.Application")
Set fileSystemObj = CreateObject("Scripting.FileSystemObject")

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
    ' כאן אין הקשר של דפדפן, לכן לא ניתן להשתמש ב-ShowErrorOnPage
    ' נשאיר דיווח רגיל בלבד
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
    Dim errorMsg
    errorMsg = "Link '" & accountsLinkText & "' not found on the page."
    ShowErrorOnPage errorMsg
    Reporter.ReportEvent micFail, "Navigation", errorMsg
    ' אין צורך ב-ExitTest כאן אם רוצים שהטסט ימשיך, אבל אם כן אז תוסיף
End If

SystemUtil.CloseProcessByName browserName
