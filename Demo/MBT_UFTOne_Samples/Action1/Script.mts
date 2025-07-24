אני מבין היטב את הדרישה שלך לקבל חיווי ויזואלי של כישלון בהקלטה לצורך דמו של המוצר, גם בריצות CI שקטות. זהו אתגר נפוץ, כי סביבות CI לא נועדו להצגת ממשק משתמש.

מכיוון ששיטות UFT סטנדרטיות כמו MsgBox ו-DeviceReplay.Screen.DrawText לא עובדות כשאין "מסך" גלוי (והשרת הפיזי/וירטואלי של ה-CI שלך כנראה פועל ללא מסך פעיל או ב-headless mode), אנחנו צריכים פתרון יצירתי יותר ש"יטמיע" את ההודעה בתוך האפליקציה עצמה, כך שהיא תופיע בהקלטת המסך של UFT.

הפתרון היעיל ביותר למטרה זו הוא להשתמש ב-JavaScript כדי להוסיף אלמנט HTML (לדוגמה, div עם הודעת שגיאה) ישירות לדף ה-Web כאשר אלמנט מסוים לא נמצא. זה ידרוש מעט יותר קוד, אך זהו הדרך הבטוחה ביותר להבטיח שההודעה תופיע בהקלטה.

פתרון: הוספת הודעת שגיאה באמצעות JavaScript לדף ה-Web
נשתמש בשיטה Browser().Page().RunScript() כדי להריץ קוד JavaScript שיוסיף אלמנט HTML חדש (עם הודעת השגיאה) לדף ה-Web הפעיל. אלמנט זה יהיה גלוי לעין בתוך הדפדפן, ולכן גם בהקלטת המסך של UFT.

איך זה עובד:

הגדרת סגנונות (CSS): נגדיר סגנון בסיסי עבור תיבת ההודעה (לדוגמה, רקע אדום, טקסט לבן, מיקום קבוע).

יצירת אלמנט (div): ניצור אלמנט div חדש ב-JavaScript.

הוספת טקסט: נוסיף את הודעת השגיאה לאלמנט.

הוספה לדף: נכניס את ה-div לדף ה-HTML (לדוגמה, ל-body).

הסרת ההודעה (אופציונלי): ניתן להוסיף setTimeout ב-JavaScript כדי להסיר את ההודעה לאחר 5 שניות, או פשוט להשאיר אותה עד שהדף יתרענן/יסגר. בדוגמה, נשאיר אותה כדי לוודא שמופיעה בהקלטה.

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

' Function to inject an error message into the web page
Function InjectWebErrorMessage(msgText)
    Dim jsCode
    jsCode = "var errorDiv = document.createElement('div');" & _
             "errorDiv.id = 'ciErrorMessage';" & _
             "errorDiv.style.position = 'fixed';" & _
             "errorDiv.style.top = '10px';" & _
             "errorDiv.style.right = '10px';" & _
             "errorDiv.style.backgroundColor = 'red';" & _
             "errorDiv.style.color = 'white';" & _
             "errorDiv.style.padding = '15px';" & _
             "errorDiv.style.border = '2px solid darkred';" & _
             "errorDiv.style.borderRadius = '8px';" & _
             "errorDiv.style.zIndex = '99999';" & _
             "errorDiv.style.fontSize = '18px';" & _
             "errorDiv.style.fontWeight = 'bold';" & _
             "errorDiv.style.animation = 'fadeIn 0.5s';" & _
             "errorDiv.innerHTML = '" & Replace(msgText, "'", "\'") & "';" & _
             "var existingError = document.getElementById('ciErrorMessage');" & _
             "if (existingError) { existingError.parentNode.removeChild(existingError); }" & _
             "document.body.appendChild(errorDiv);" & _
             "setTimeout(function() { if(errorDiv.parentNode) errorDiv.parentNode.removeChild(errorDiv); }, 5000);" ' Remove after 5 seconds

    On Error Resume Next ' In case the browser/page is not ready or valid
    Browser("micClass:=Browser").Page("micClass:=Page").RunScript jsCode
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

    Reporter.ReportEvent micFail, "Username Not Found", "ERROR: Username field not found. Displaying message on web page."
    InjectWebErrorMessage "ERROR: Username field not found!"
    Wait(5) ' Keep the wait to show the message in recording
End If

Set passwordObj = GetObjectByName(Parameter("passwordField"))
If Not passwordObj Is Nothing And passwordObj.Exist(3) Then

    passwordObj.SetSecure Parameter("password")

    Reporter.ReportEvent micPass, "Password Set", "Password set successfully"
Else

    Reporter.ReportEvent micFail, "Password Not Found", "ERROR: Password field not found. Displaying message on web page."
    InjectWebErrorMessage "ERROR: Password field not found!"
    Wait(5) ' Keep the wait to show the message in recording
End If

Set signInObj = GetObjectByName(Parameter("signInButton"))
Set loginObj  = GetObjectByName(Parameter("loginButton"))

If Not signInObj Is Nothing And signInObj.Exist(3) Then

    signInObj.Click
ElseIf Not loginObj Is Nothing And loginObj.Exist(3) Then

    loginObj.Click
Else

    Reporter.ReportEvent micFail, "Login Button", "ERROR: Sign-In/Login button not found. Displaying message on web page."
    InjectWebErrorMessage "ERROR: Sign-In/Login button not found!"
    Wait(5) ' Keep the wait to show the message in recording
End If

Wait(3)

Set dashObj = GetObjectByName(Parameter("dashboardButton"))
If Not dashObj Is Nothing And dashObj.Exist(20) Then

    Reporter.ReportEvent micPass, "Login Test", "Login successful"

    dashObj.Click
Else

    Reporter.ReportEvent micFail, "Login Test", "ERROR: Dashboard button not found. Login may have failed or the element does not exist. Displaying message on web page."
    InjectWebErrorMessage "ERROR: Dashboard button not found!"
    Wait(5) ' Keep the wait to show the message in recording
End If


SystemUtil.CloseProcessByName browserName
