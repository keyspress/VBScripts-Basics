option explicit

dim a
dim b
dim c
dim length
dim obj

a = wscript.scriptname
b = wscript.scriptfullname

msgbox a
msgbox b

length = len(b) - len(a)
c = left(b, length)

msgbox c

set obj = createobject("wscript.shell")

msgbox obj.CurrentDirectory