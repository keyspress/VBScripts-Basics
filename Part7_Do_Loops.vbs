option explicit

dim a
dim password

a = 0

do until a = 5
a = a + 1
msgbox a
loop


msgbox"sup from outside the loop"

do 
a = a + 1
msgbox a
loop while a < 8

msgbox"You made it out of the second loop"

do 
password = inputbox("Password")
loop until password = "1234"

msgbox"How did you guess?!"