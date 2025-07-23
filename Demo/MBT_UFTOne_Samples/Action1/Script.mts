' ������️ קבלת שם הלינק לפרמטר (ברירת מחדל: "Accounts")
Dim accountsLinkText
accountsLinkText = Trim(Parameter("ElementName"))
If accountsLinkText = "" Then accountsLinkText = "Accounts"

' ������ שימוש בזיהוי דינמי ללינק (במקום מתוך OR)
Dim linkDesc
Set linkDesc = Description.Create()
linkDesc("micclass").Value = "Link"
linkDesc("innertext").Value = accountsLinkText

If Browser("Dashboard - Advantage").Page("Dashboard - Advantage").ChildObjects(linkDesc).Count > 0 Then
    Wait(3)
    Browser("Dashboard - Advantage").Page("Dashboard - Advantage").ChildObjects(linkDesc)(0).Click
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
    Reporter.ReportEvent micFail, "Navigation", "Link '" & accountsLinkText & "' not found on dashboard"
    ExitTest
End If

