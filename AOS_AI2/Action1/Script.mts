Dim iURL

Dim objShell

iURL = "https://advantageonlineshopping.com/#/"

set objShell = CreateObject("Shell.Application")

Set fileSystemObj = createobject("Scripting.FileSystemObject")

chromeExist = "C:\Program Files\Google\Chrome\Application\chrome.exe"

If fileSystemObj.FileExists(chromeExist) then

objShell.ShellExecute "C:\Program Files\Google\Chrome\Application\chrome.exe", iURL, "", ""

Else

objShell.ShellExecute "C:\Program Files\Google\Chrome\Application\chrome.exe", iURL, "", ""

End If

AIUtil.SetContext Window("regexpwndtitle:=Google Chrome", "regexpwndclass:=Chrome_WidgetWin_1", "is owned window:=False", "is child window:=False")
AIUtil("profile", micAnyText, micFromBottom, 1).Click
AIUtil("input", "Username").SetText "demorun1"
AIUtil("input", "Password").SetText "Password1"
AIUtil("button", "SIGN IN").Click
