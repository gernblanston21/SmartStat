' Safe logging pattern snippet (no MsgBox; operator notified via gui:error_message)
Option Explicit

Sub LogErrorExample(ByVal logPath, ByVal msg)
    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set ts = fso.OpenTextFile(logPath, 8, True) ' ForAppending
    If Err.Number = 0 Then
        ts.WriteLine Now() & " | ERROR | " & msg
        ts.Close
    End If
    On Error GoTo 0

    Call TrioCmd("gui:error_message " & Chr(34) & "Error logged: " & logPath & Chr(34))
End Sub
