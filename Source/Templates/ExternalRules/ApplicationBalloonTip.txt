
Line1 = "Did you know you can geneerate balloon tips with iLogic rules?"
Line2 = "MFG601910, Bridging the Gap Between iLogic Automation and Inventor Add-Ins" 

oMsg = Line1 & vbLf & vbLf & Line2

oTipNumber = Int((99 * Rnd) + 1) ' Random number between 1 and 99.
oName = "Balloon Tip " & oTipNumber

Dim oBtip As BalloonTip
Try
	oBtip = ThisApplication.UserInterfaceManager.BalloonTips.Add("_Tip" & oTipNumber , oName , oMsg)
Catch
	oBtip = ThisApplication.UserInterfaceManager.BalloonTips.Item("_Tip" & oTipNumber)	
End Try

oBtip.Message = oMsg
oBtip.Interval = 500
oBtip.Display
