' Viz Trio tabfield property examples (repo snippet)
' NOTE: Actual TrioCmd availability/behavior must be grounded in docs/viz-trio.

Option Explicit

Sub Example_GetSet()
    Dim tabName: tabName = "H0010"
    Dim val

    ' Get
    val = TrioCmd("page:get_property " & tabName)

    ' Set
    Call TrioCmd("page:set_property " & tabName & " " & Quote("Jordan"))

    ' Example of safe operator notification pattern (no blocking MsgBox)
    Call TrioCmd("gui:error_message " & Quote("Example completed. Value was: " & val))
End Sub

Private Function Quote(ByVal s)
    Quote = Chr(34) & s & Chr(34)
End Function
