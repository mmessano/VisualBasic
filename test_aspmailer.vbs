Set Mailer = CreateObject("SMTPsvg.Mailer")
Mailer.FromName   = "Joe’s Widgets Corp."
Mailer.FromAddress= "mmessano@dexma.com"
Mailer.RemoteHost = "outbound.smtp.dexma.com"
'Mailer.AddRecipient "John Smith", "jsmith@anotherhostname.com"
Mailer.Subject    = "Great SMTP Product!"
Mailer.BodyText   = "Dear Stephen" & VbCrLf & "Your widgets order has been processed!"
if Mailer.SendMail then
  Response.Write "Mail sent..."
else
  Response.Write "Mail send failure. Error was " & Mailer.Response
end if