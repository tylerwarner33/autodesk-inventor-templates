
Dim activeTheme As String = ThisApplication.ThemeManager.ActiveTheme.Name

If activeTheme = "DarkTheme"  Then
	MsgBox("You're rockin the dark theme",, "iLogic")	
ElseIf activeTheme = "LightTheme" Then
	MsgBox("You're rockin the light theme",, "iLogic")	
End If
