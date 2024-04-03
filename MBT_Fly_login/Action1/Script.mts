systemutil.Run("C:\Program Files (x86)\Micro Focus\UFT One\samples\Flights Application\FlightsGUI.exe")

'login

WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "john" @@ hightlight id_;_2102821832_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "660ce8121029053eaf54" @@ hightlight id_;_1950641056_;_script infofile_;_ZIP::ssf4.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_1950642304_;_script infofile_;_ZIP::ssf3.xml_;_
