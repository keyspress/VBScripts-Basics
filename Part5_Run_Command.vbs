Option Explicit

Dim obj, a, b, c

set obj = CreateObject("wscript.shell")

a = MsgBox("Wanna Open Some Shit Motherfucker?", vbYesNo+vbQuestion+vbSystemModal)
if a = vbYes then 
   obj.run "notepad.exe"
   b = MsgBox("What about a Folder? Do you want that?", vbYesNo+vbQuestion+vbSystemModal)
else
   b = MsgBox("Ugh..Fine... What about a Folder? Do you want that?", vbYesNo+vbQuestion+vbSystemModal)   
end if  

if b = vbYes then
   obj.run "C:\Users\Kyle\Dropbox\Guitar\Arpeggios"
   c = MsgBox("Now do you want to really fuck some shit up?", vbYesNo+vbQuestion+vbSystemModal)
else   
   c = MsgBox("Whatever, your loss. So do you wannaa really fuck some shit up?", vbYesNo+vbQuestion+vbSystemModal)
end if
if c = vbYes then   
   obj.run "powershell.exe"
else
   WScript.Quit
end if      