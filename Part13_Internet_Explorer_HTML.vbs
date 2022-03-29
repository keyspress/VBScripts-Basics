
On Error Resume Next

ie.Document.All.Item("email").Value = "somerandomcrap@test.com"
ie.Document.All.Item("pass").Value = "apassword"

ie.Document.All.Item("login_form").Submit