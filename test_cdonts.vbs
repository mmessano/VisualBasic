'Send the error report using email. This currently uses CDONTS, but could be
'adapted to use another mail sending object if required
    Set myMail = CreateObject("CDONTS.NewMail")

    With myMail
        .From = MAIL_FROM_NAME & "<" & MAIL_FROM_EMAIL & ">"
        .To = MAIL_TO_NAME & "<" & MAIL_TO_EMAIL & ">"
        .Subject = MAIL_SUBJECT
        .Value("MIME-Version") = "1.1"
        .BodyFormat=0
        .MailFormat=0
        .Body=HTML
        .Send
    End With

    Set myMail = Nothing