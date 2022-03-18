option explicit

dim ie, x
set x = CreateObject("wscript.shell")
set ie = CreateObject("InternetExplorer.Application")

ie.Navigate"https://www.cnn.com/"
ie.Visible=true
ie.Toolbar=0
ie.StatusBar=false
ie.Height=560
ie.Width=1000
ie.Top=0
ie.Left=0
ie.Resizeable=0
ie.Visible=1

Do while ie.Busy
  wscript.sleep 200
Loop  
