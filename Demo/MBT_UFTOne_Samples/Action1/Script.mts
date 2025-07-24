אני מבין לחלוטין את הצורך שלך בפתרון ודאי שיציג את הודעת השגיאה בהקלטה של UFT בריצות CI שקטות, במיוחד לדמו של המוצר. הניסיונות הקודמים עם MsgBox ו-DeviceReplay.Screen.DrawText נכשלו מכיוון שסביבות CI לעיתים קרובות פועלות ללא ממשק משתמש גלוי או ב-headless mode.

הפתרון הוודאי ביותר במצב זה הוא לגרום לאפליקציית ה-Web עצמה להציג את הודעת השגיאה. כלומר, אנחנו נטמיע את הודעת השגיאה ישירות בתוך קוד ה-JavaScript שרץ בדפדפן, כך שהיא תהפוך לחלק מה-DOM של העמוד ותהיה גלויה ל-UFT (ולכן גם בהקלטה).

איך זה יעבוד בוודאות?

UFT מקליט את מה שהדפדפן מציג. אם הדפדפן עצמו (בזכות קוד ה-JavaScript שנדחוף אליו) מציג את ההודעה, UFT יקליט אותה. זה לא תלוי ביכולות ציור חיצוניות של UFT או בחלונות קופצים של מערכת ההפעלה.

פתרון בטוח: הזרקת הודעת שגיאה ל-DOM של דף ה-Web באמצעות JavaScript
הפתרון הזה כולל הוספת פונקציה שתשתמש ב-JavaScript כדי:

ליצור אלמנט HTML חדש (לדוגמה, div).

לעצב אותו בסגנון בולט (רקע אדום, טקסט גדול וכו').

להוסיף לו את הודעת השגיאה.

להכניס אותו לגוף הדף (body), מה שהופך אותו לחלק מהתוכן הגלוי של הדפדפן.

להסיר אותו אוטומטית לאחר 5 שניות.

סקריפט 1 מעודכן: כניסה למערכת (Login)
קטע קוד

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

    Reporter.ReportEvent micFail, "Browser Launch", "No supported browser found"

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


Function GetObjectByName(elementName)

    Select Case elementName

        Case "username"

            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("username")

        Case "password"

            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("password")

        Case "signIn"

            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Sign-In")

        Case "login"

            Set GetObjectByName = Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Login")

        Case "dashboardBtn"

            Set GetObjectByName = Browser("Dashboard - Advantage_2").Page("Dashboard - Advantage").WebElement("Bank Accounts")

        Case Else

            Set GetObjectByName = Nothing

    End Select
End Function

Set usernameObj = GetObjectByName(Parameter("usernameField"))
If Not usernameObj Is Nothing And usernameObj.Exist(3) Then
    usernameObj.Set Parameter("username")
    Reporter.ReportEvent micPass, "Username Set", "Username set successfully"
Else
    Reporter.ReportEvent micFail, "Username Not Found", "ERROR: Username field not found. Injected message to web page."
    InjectWebErrorMessage "ERROR: Username field '" & Parameter("usernameField") & "' not found!"
    Wait(5) ' Keep the wait to ensure message is visible in recording
End If

Set passwordObj = GetObjectByName(Parameter("passwordField"))
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then
    passwordObj.SetSecure Parameter("password")
    Reporter.ReportEvent micPass, "Password Set", "Password set successfully"
Else
    Reporter.ReportEvent micFail, "Password Not Found", "ERROR: Password field not found. Injected message to web page."
    InjectWebErrorMessage "ERROR: Password field '" & Parameter("passwordField") & "' not found!"
    Wait(5) ' Keep the wait to ensure message is visible in recording
End If

Set signInObj = GetObjectByName(Parameter("signInButton"))
Set loginObj  = GetObjectByName(Parameter("loginButton"))

If Not signInObj Is Nothing And signInObj.Exist(3) Then
    signInObj.Click
ElseIf Not loginObj Is Nothing And loginObj.Exist(3) Then
    loginObj.Click
Else
    Reporter.ReportEvent micFail, "Login Button", "ERROR: Sign-In/Login button not found. Injected message to web page."
    InjectWebErrorMessage "ERROR: Login button ('" & Parameter("signInButton") & "' or '" & Parameter("loginButton") & "') not found!"
    Wait(5) ' Keep the wait to ensure message is visible in recording
End If

Wait(3)

Set dashObj = GetObjectByName(Parameter("dashboardButton"))
If Not dashObj Is Nothing And dashObj.Exist(20) Then
    Reporter.ReportEvent micPass, "Login Test", "Login successful"
    dashObj.Click
Else
    Reporter.ReportEvent micFail, "Login Test", "ERROR: Dashboard button not found. Login may have failed or the element does not exist. Injected message to web page."
    InjectWebErrorMessage "ERROR: Dashboard button ('" & Parameter("dashboardButton") & "') not found!"
    Wait(5) ' Keep the wait to ensure message is visible in recording
End If


SystemUtil.CloseProcessByName browserName
