
Set objShell = WScript.CreateObject( "WScript.Shell" )
strName = objShell.ExpandEnvironmentStrings("%USERNAME%")
objShell.Run("""C:\Users\" & strName & "\AppData\Roaming\Spotify\Spotify.exe""")
Set oShell = CreateObject("WScript.Shell")
Result = oShell.AppActivate("Spotify Premium")
WScript.Sleep 3000
oShell.SendKeys "+"
WScript.Sleep 10
oShell.SendKeys "+t"
oShell.SendKeys "+"
WScript.Sleep 10
oShell.SendKeys "+h"
WScript.Sleep 10
oShell.SendKeys "+x"
WScript.Sleep 10
oShell.SendKeys "1"
WScript.Sleep 10
oShell.SendKeys "1"
WScript.Sleep 10
oShell.SendKeys "3"
WScript.Sleep 10
oShell.SendKeys "8"
