Option Explicit

dim birth

birth = msgbox("Is it your birthday?", vbYesNo+vbQuestion, "Tell me")

if birth = vbYes then 
  msgbox "Happy Birthday Mother Fucker!", vbInformation
else 
  msgbox "Well that sucks, no cake for you!"
end if