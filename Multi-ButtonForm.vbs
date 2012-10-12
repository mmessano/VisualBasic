'LISTING 2: Multi-ButtonForm.vbs
Option Explicit
Dim gsForm
Set gsForm = WScript.CreateObject("GooeyScript.Form", "gsForm_")
gsForm.Load ,,,,"This is a multi-button form",,,FALSE
gsForm.Button.Load "Button1",100,50,100,50,"Button 1", TRUE
gsForm.Button.Load "Button2",350,50,100,50,"Button 2", TRUE
gsForm.Button.Load "Button3",100,150,100,50,"Button 3", TRUE
gsForm.Button.Load "Button4",350,150,100,50,"Button 4", TRUE
gsForm.OnTop = TRUE
gsForm.Visible = TRUE
gsForm.Pause

'BEGIN COMMENT LINE
'Execution pauses now and recommences here after Form::UnPause has been
'called.
'END COMMENT LINE
gsForm.Unload
Set gsForm = Nothing

Sub gsForm_ButtonClick(strButtonName)
	BEGIN CALLOUT A
		Select Case strButtonName
			case "Button2" : HideButton1
			case "Button3" : SayHi3
			case "Button4" : CloseMyForm
		End Select
	END CALLOUT A
End Sub

Sub gsForm_FormClose()
	gsForm.OnTop = TRUE
	gsForm.Visible = TRUE
End Sub

Sub HideButton1
	gsForm.Button.Visible "Button1", FALSE
End Sub

Sub SayHi3
	gsForm.OnTop = FALSE
	MsgBox "Hi! Button 3 pressed"
	gsForm.OnTop = TRUE
End Sub

Sub CloseMyForm
	gsForm.UnPause
End Sub