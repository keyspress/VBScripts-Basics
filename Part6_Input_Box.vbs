' InputBox "So Watcha Watcha Watcha Name?", "Yo Name:", "Put yo name here.",13000,9000
Option Explicit

dim name
dim nameMsg

dim quest
dim questMsg

dim airSpeed
dim airSpeedMsg

name = InputBox("What, is your name?", "Name:")
nameMsg = "Hi your name is what? Your name is who? Your name is shika shika " + name

msgbox nameMsg, vbOKOnly, "This is your name."

quest = InputBox("What, is your quest?", "Quest")
questMsg = "Alright, fair enough. Good luck with " + quest

msgbox questMsg, vbOKOnly, "Quest"

airSpeed = InputBox("What, is the airspeed velocity of an unladen swallow?", "Airspeed")
if airSpeed = "I don't know" then
  airSpeedMsg = "You are thrown into a deep chasm never to be seen again. Because you are dead. At the bottom of the chasm. Being eaten by bugs. So IDK, I guess they see you at least"
  msgbox airSpeedMsg, vbOKOnly, "The End"
elseif airSpeed = "What do you mean, an african or european swallow?" then
  airSpeedMsg = "What? I don't know thaaaaaagghhhaahhhh!!!!!!!"
  msgbox airSpeedMsg, vbOKOnly, "The End"
elseif airSpeed <> "fksdfj;asdlkfjsdklfhnasdiklfasdhn;lasd" then
  airSpeedMsg = "Ni!!!!!!" 
  msgbox airSpeedMsg, vbOKOnly, "The End"
end if 
