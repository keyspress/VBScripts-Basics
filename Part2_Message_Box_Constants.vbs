a = msgbox ("Panda Panda Panda",_
vbAbortRetryIgnore+vbQuestion+vbDefaultButton2+vbSystemModal,_
"Gadget")

if a = vbAbort then msgbox "Fine, fuck off!", vbCritical
