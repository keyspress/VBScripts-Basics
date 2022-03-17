

set x = createobject("wscript.shell")

x.run "notepad.exe"

wscript.sleep 1000

x.sendkeys "OOOOOO I am the ghost in the shell oooooOOOOOO I am the ghost in the shell ooooo"
x.sendkeys "{enter}"
wscript.sleep 800
x.sendkeys "I'm just kidding.....or am I? oooooooo"
wscript.sleep 600
x.sendkeys "%fs"
x.sendkeys "test.vbs"
wscript 300
' x.sendkeys"{enter}"
' x.sendkeys "{tab}" doesn't work
' x.sendkeys "{tab}"
' x.sendkeys "{tab}"
' x.sendkeys "{tab}"