Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Forms
Imports Inventor

Module Welcome

    'This command has no button, it runs on an event

    'This is the code that does the real work when your command is executed.
    Sub Addin_Hello(DocName As String, EventName As String)

        If DocName = "" Then DocName = "a new file"

        EventName = Replace(EventName, "k", "")
        Dim UserName = Environ$("UserName")
        Dim AddinName = Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        MsgBox("My Name is " & AddinName & " , I'm an add-in." _
               & vbLf & "I fired on the " & EventName & " event of:" & vbLf & DocName,, "Hello " & UserName)

    End Sub

End Module
