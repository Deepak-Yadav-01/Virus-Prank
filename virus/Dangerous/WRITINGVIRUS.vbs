do
Set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "H"
wscript.sleep 100
wshshell.sendkeys "A"
wscript.sleep 100
wshshell.sendkeys "C"
wscript.sleep 100
wshshell.sendkeys "K"
wscript.sleep 100
wshshell.sendkeys "E"
wscript.sleep 100
wshshell.sendkeys "D"
wscript.sleep 100
wshshell.sendkeys "  "
wscript.sleep 100
wshshell.sendkeys "!"
loop