Imports Inventor
Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms

Module Welcome

    ' This command has no button, it runs on an event.

    ' This is the code that does the real work when your command is executed.
    Sub AddinHello(documentName As String, eventName As String)

        If documentName = "" Then documentName = "a new file"

        eventName = Replace(eventName, "k", "")
        Dim userName = Environ$("UserName")
        Dim addinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        MsgBox("My Name is " & addinName & " , I'm an add-in." _
            & vbLf & "I fired on the " & eventName & " event of:" & vbLf & documentName,, "Hello " & userName)

    End Sub

End Module
