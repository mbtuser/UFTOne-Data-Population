Dim iURL
Dim objShell
iURL = "https://advantageonlinebanking.com/dashboard"
set objShell = CreateObject("Shell.Application")

Set fileSystemObj = createobject("Scripting.FileSystemObject")
edgeExist = "C:\Program Files\Mozilla Firefox\firefox.exe"
If fileSystemObj.FileExists(edgeExist) then
objShell.ShellExecute "C:\Program Files\Mozilla Firefox\firefox.exe", iURL, "", ""
Else
objShell.ShellExecute "C:\Program Files (x86)\Mozilla Firefox\firefox.exe", iURL, "", ""
End If
wait(3)
If Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Exist(3) Then
	Browser("Browser").Page("Dashboard - Advantage").WebButton("WebButton").Click
       Browser("Browser").Page("Dashboard - Advantage").WebMenu("My Profile Management").Select "Logout"
End If
wait(5) @@ script infofile_;_ZIP::ssf18.xml_;_
	Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("username").Set Parameter("username")
	Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebEdit("password").SetSecure Parameter("password")
	If Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Sign-In").Exist Then
		Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Sign-In").Click
		Else
		Browser("Home - Advantage Bank").Page("Home - Advantage Bank").WebButton("Login").Click
	End If
wait(3)
  If Browser("Home - Advantage Bank").Page("Dashboard - Advantage").WebButton("WebButton").Exist(5) Then
 	Reporter.ReportEvent micPass, "Passed Login Test", "Login succefull"
 	else
 	Reporter.ReportEvent micFail, "Failed", "Fail to Login incurrect user or password"
 End If
 systemUtil.CloseProcessByName ("firefox.exe")
