On Error Resume Next
Set conn = CreateObject("ADODB.Connection")
If Err.Number <> 0 Then
    WScript.Echo "ADO not registered: " & Err.Description
Else
    WScript.Echo "ADO connection object created successfully."
End If
