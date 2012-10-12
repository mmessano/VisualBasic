Dim userMail
Set userMail = Server.CreateObject("CDONTS.NewMail")
userMail.To = "mmessano@dexma.com"
userMail.From = "mmessano@dexma.com"
userMail.Subject = "test"
userMail.Body = "test"
userMail.Send
Set userMail = Nothing