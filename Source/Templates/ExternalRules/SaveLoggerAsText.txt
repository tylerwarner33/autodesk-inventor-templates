Class ThisRule

    Shared oStartMarker As String
    Shared oStartTime As DateTime
    Shared oElapsedTime As TimeSpan
    Shared oTimeout As Integer

    Sub Main()

        Call StartLogging()

        Logger.info("")
        Logger.info("Did you know you can output logger text with an iLogic rule?")
        Logger.info("MFG601910, Bridging the Gap Between iLogic Automation and Inventor Add-Ins")
        Logger.info("")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")
        Logger.info("Hello World!")

        Call EndLogging()

        ThisApplication.UserInterfaceManager.DockableWindows.Item("ilogic.logwindow").Visible = True

    End Sub

    Sub StartLogging()

        oStartMarker = "ilogic Log"
        Logger.Info(oStartMarker)
        Logger.Info(DateTime.Now)

        oStartTime = Now

    End Sub

    Sub EndLogging()

        oElapsedTime = Now().Subtract(oStartTime)
        oTime = oElapsedTime.TotalSeconds
        oTime = Round(oTime, 2)
        oLine1 = "iLogic Complete! "
        oLine2 = "Total Run Time : " & oTime & " seconds "
        oTimeout = 3.5 ' seconds

        Logger.Info(oLine1)
        Logger.Info(oLine2)

        ' Write out the log file .
        Dim oLogFileName = "C:\Temp\iLogic Log.txt"
        iLogicVb.Automation.LogControl.SaveLogAs(oLogFileName)

        Dim oLogText = System.IO.File.ReadAllText(oLogFileName)
        Dim oMarkerIndex = oLogText.LastIndexOf(oStartMarker)

        ' Replace the file contents with the portion from the last marker.
        System.IO.File.WriteAllText(oLogFileName, oLogText.Substring(oMarkerIndex))

        ThisDoc.Launch(oLogFileName)

    End Sub

End Class
