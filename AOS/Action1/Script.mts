
Dim iURL

Dim objShell

iURL = "https://www.advantageonlineshopping.com/"

set objShell = CreateObject("Shell.Application")

Set fileSystemObj = createobject("Scripting.FileSystemObject")

chromeExist = "C:\Program Files\Mozilla Firefox\firefox.exe"

If fileSystemObj.FileExists(chromeExist) then

objShell.ShellExecute "C:\Program Files\Mozilla Firefox\firefox.exe", iURL, "", ""

Else

objShell.ShellExecute "C:\Program Files (x86)\Mozilla Firefox\firefox.exe", iURL, "", ""

End If

Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("username").Set "demorun1" @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("password").SetSecure "65e6e0ee05267349c0e36acdcbdea69bc462f902ec550dba" @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("sign_in_btn").Click @@ script infofile_;_ZIP::ssf4.xml_;_
