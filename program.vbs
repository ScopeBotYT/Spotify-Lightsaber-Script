Set oShell = CreateObject("WScript.Shell")
oShell.Run("""C:\Users\Janni\AppData\Roaming\Spotify\Spotify.exe""")
WScript.Sleep 3000
oShell.SendKeys "+"
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