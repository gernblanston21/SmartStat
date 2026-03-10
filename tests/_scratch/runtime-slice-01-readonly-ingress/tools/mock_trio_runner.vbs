Option Explicit

Dim gFixture
Dim gPage
Dim gCustom
Dim gAllCalls
Dim gMutationCalls
Dim gUnknownCalls

Set gFixture = CreateObject("Scripting.Dictionary")
Set gPage = CreateObject("Scripting.Dictionary")
Set gCustom = CreateObject("Scripting.Dictionary")
Set gAllCalls = CreateObject("System.Collections.ArrayList")
Set gMutationCalls = CreateObject("System.Collections.ArrayList")
Set gUnknownCalls = CreateObject("System.Collections.ArrayList")

Main

Sub Main()
  Dim targetScript, fixturePath, reportOut
  targetScript = GetNamedArg("target_script", "")
  fixturePath = GetNamedArg("fixture", "")
  reportOut = GetNamedArg("report_out", "")

  If Len(targetScript) = 0 Or Len(fixturePath) = 0 Or Len(reportOut) = 0 Then
    WScript.Echo "mock_trio_runner usage: /target_script:<path> /fixture:<path> /report_out:<path>"
    WScript.Quit 2
  End If

  Call LoadFixture(fixturePath)
  Call ExecuteTargetScript(targetScript)
  Call WriteReport(reportOut, targetScript)
End Sub

Function TrioCmd(ByVal commandText)
  Dim normCmd
  normCmd = NormalizeCommand(commandText)
  gAllCalls.Add normCmd

  If Left(normCmd, 18) = "PAGE:SET_PROPERTY " Then
    Call HandlePageSet(normCmd)
    TrioCmd = "1"
    Exit Function
  End If

  If Left(normCmd, 29) = "TABFIELD:SET_CUSTOM_PROPERTY " Then
    Call HandleCustomSet(normCmd)
    TrioCmd = "1"
    Exit Function
  End If

  If Left(normCmd, 22) = "SOCK:SEND_SOCKET_DATA " Then
    gMutationCalls.Add normCmd
    TrioCmd = "1"
    Exit Function
  End If

  If Left(normCmd, 17) = "GUI:ERROR_MESSAGE " Then
    gMutationCalls.Add normCmd
    TrioCmd = "1"
    Exit Function
  End If

  If normCmd = "SOCK:SOCKET_IS_CONNECTED" Then
    TrioCmd = LookupFixtureValue(normCmd, "0")
    Exit Function
  End If

  If Left(normCmd, 18) = "PAGE:GET_PROPERTY " Then
    TrioCmd = HandlePageGet(normCmd)
    Exit Function
  End If

  If Left(normCmd, 29) = "TABFIELD:GET_CUSTOM_PROPERTY " Then
    TrioCmd = HandleCustomGet(normCmd)
    Exit Function
  End If

  If gFixture.Exists(normCmd) Then
    TrioCmd = CStr(gFixture(normCmd))
    Exit Function
  End If

  gUnknownCalls.Add normCmd
  TrioCmd = ""
End Function

Sub HandlePageSet(ByVal normCmd)
  Dim rest, sp, tabName, valueText
  rest = Mid(normCmd, 19)
  sp = InStr(rest, " ")
  If sp <= 1 Then
    gMutationCalls.Add normCmd
    Exit Sub
  End If

  tabName = Trim(Left(rest, sp - 1))
  valueText = Trim(Mid(rest, sp + 1))
  valueText = StripOuterQuotes(valueText)

  gPage(tabName) = valueText
  gMutationCalls.Add "PAGE:SET_PROPERTY " & tabName & "=" & valueText
End Sub

Sub HandleCustomSet(ByVal normCmd)
  Dim rest, sp, tabName, valueText
  rest = Mid(normCmd, 30)
  sp = InStr(rest, " ")
  If sp <= 1 Then
    gMutationCalls.Add normCmd
    Exit Sub
  End If

  tabName = Trim(Left(rest, sp - 1))
  valueText = Trim(Mid(rest, sp + 1))
  valueText = StripOuterQuotes(valueText)

  gCustom(tabName) = valueText
  gMutationCalls.Add "TABFIELD:SET_CUSTOM_PROPERTY " & tabName & "=" & valueText
End Sub

Function HandlePageGet(ByVal normCmd)
  Dim tabName
  tabName = Trim(Mid(normCmd, 19))
  If gPage.Exists(tabName) Then
    HandlePageGet = CStr(gPage(tabName))
    Exit Function
  End If
  If gFixture.Exists(normCmd) Then
    HandlePageGet = CStr(gFixture(normCmd))
    Exit Function
  End If
  HandlePageGet = ""
End Function

Function HandleCustomGet(ByVal normCmd)
  Dim tabName
  tabName = Trim(Mid(normCmd, 30))
  If gCustom.Exists(tabName) Then
    HandleCustomGet = CStr(gCustom(tabName))
    Exit Function
  End If
  If gFixture.Exists(normCmd) Then
    HandleCustomGet = CStr(gFixture(normCmd))
    Exit Function
  End If
  HandleCustomGet = ""
End Function

Sub LoadFixture(ByVal fixturePath)
  Dim fso, ts, line, eqPos, k, v
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(fixturePath) Then
    WScript.Echo "fixture not found: " & fixturePath
    WScript.Quit 3
  End If

  Set ts = fso.OpenTextFile(fixturePath, 1, False)
  Do Until ts.AtEndOfStream
    line = Trim(CStr(ts.ReadLine))
    If Len(line) = 0 Then
    ElseIf Left(line, 1) = "#" Then
    Else
      eqPos = InStr(1, line, "=", vbBinaryCompare)
      If eqPos > 1 Then
        k = NormalizeCommand(Left(line, eqPos - 1))
        v = Mid(line, eqPos + 1)
        gFixture(k) = v
        Call SeedStateFromFixture(k, v)
      End If
    End If
  Loop
  ts.Close
End Sub

Sub SeedStateFromFixture(ByVal fixtureKey, ByVal fixtureValue)
  Dim tabName
  If Left(fixtureKey, 18) = "PAGE:GET_PROPERTY " Then
    tabName = Trim(Mid(fixtureKey, 19))
    gPage(tabName) = CStr(fixtureValue)
    Exit Sub
  End If

  If Left(fixtureKey, 29) = "TABFIELD:GET_CUSTOM_PROPERTY " Then
    tabName = Trim(Mid(fixtureKey, 30))
    gCustom(tabName) = CStr(fixtureValue)
    Exit Sub
  End If
End Sub

Sub ExecuteTargetScript(ByVal targetScript)
  Dim fso, ts, codeText
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(targetScript) Then
    WScript.Echo "target script not found: " & targetScript
    WScript.Quit 4
  End If

  Set ts = fso.OpenTextFile(targetScript, 1, False)
  codeText = ts.ReadAll
  ts.Close

  ExecuteGlobal codeText
End Sub

Sub WriteReport(ByVal reportPath, ByVal targetScript)
  Dim fso, ts, keyList, i
  Dim headerLine, versionConst

  Set fso = CreateObject("Scripting.FileSystemObject")
  Call EnsureParentFolder(reportPath)
  Set ts = fso.CreateTextFile(reportPath, True, False)

  headerLine = ReadFirstLine(targetScript)
  versionConst = ReadVersionConst(targetScript)

  ts.WriteLine "TARGET_SCRIPT=" & targetScript
  ts.WriteLine "SCRIPT_HEADER=" & headerLine
  ts.WriteLine "SCRIPT_VERSION_CONST=" & versionConst
  ts.WriteLine "TOTAL_CALLS=" & CStr(gAllCalls.Count)
  ts.WriteLine "TOTAL_MUTATION_CALLS=" & CStr(gMutationCalls.Count)
  ts.WriteLine "TOTAL_UNKNOWN_CALLS=" & CStr(gUnknownCalls.Count)

  ts.WriteLine "MUTATION_CALLS_BEGIN"
  For i = 0 To gMutationCalls.Count - 1
    ts.WriteLine CStr(gMutationCalls(i))
  Next
  ts.WriteLine "MUTATION_CALLS_END"

  ts.WriteLine "PAGE_STATE_BEGIN"
  keyList = SortedKeys(gPage)
  For i = LBound(keyList) To UBound(keyList)
    ts.WriteLine CStr(keyList(i)) & "=" & CStr(gPage(CStr(keyList(i))))
  Next
  ts.WriteLine "PAGE_STATE_END"

  ts.WriteLine "CUSTOM_STATE_BEGIN"
  keyList = SortedKeys(gCustom)
  For i = LBound(keyList) To UBound(keyList)
    ts.WriteLine CStr(keyList(i)) & "=" & CStr(gCustom(CStr(keyList(i))))
  Next
  ts.WriteLine "CUSTOM_STATE_END"

  ts.WriteLine "UNKNOWN_CALLS_BEGIN"
  For i = 0 To gUnknownCalls.Count - 1
    ts.WriteLine CStr(gUnknownCalls(i))
  Next
  ts.WriteLine "UNKNOWN_CALLS_END"

  ts.Close
End Sub

Function LookupFixtureValue(ByVal fixtureKey, ByVal defaultValue)
  If gFixture.Exists(fixtureKey) Then
    LookupFixtureValue = CStr(gFixture(fixtureKey))
  Else
    LookupFixtureValue = CStr(defaultValue)
  End If
End Function

Function SortedKeys(ByVal d)
  Dim keys, i, j, tmp
  If d.Count = 0 Then
    SortedKeys = Array()
    Exit Function
  End If

  keys = d.Keys
  For i = LBound(keys) To UBound(keys) - 1
    For j = i + 1 To UBound(keys)
      If StrComp(CStr(keys(j)), CStr(keys(i)), vbBinaryCompare) < 0 Then
        tmp = keys(i)
        keys(i) = keys(j)
        keys(j) = tmp
      End If
    Next
  Next
  SortedKeys = keys
End Function

Function NormalizeCommand(ByVal cmdText)
  Dim s
  s = UCase(Trim(CStr(cmdText)))
  Do While InStr(s, "  ") > 0
    s = Replace(s, "  ", " ")
  Loop
  NormalizeCommand = s
End Function

Function StripOuterQuotes(ByVal txt)
  Dim s
  s = Trim(CStr(txt))
  If Len(s) >= 2 Then
    If Left(s, 1) = """" And Right(s, 1) = """" Then
      StripOuterQuotes = Mid(s, 2, Len(s) - 2)
      Exit Function
    End If
  End If
  StripOuterQuotes = s
End Function

Function ReadFirstLine(ByVal filePath)
  Dim fso, ts
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.OpenTextFile(filePath, 1, False)
  If ts.AtEndOfStream Then
    ReadFirstLine = ""
  Else
    ReadFirstLine = ts.ReadLine
  End If
  ts.Close
End Function

Function ReadVersionConst(ByVal filePath)
  Dim fso, ts, line, eqPos
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.OpenTextFile(filePath, 1, False)
  ReadVersionConst = ""
  Do Until ts.AtEndOfStream
    line = Trim(CStr(ts.ReadLine))
    If Left(UCase(line), Len("CONST SMARTSTAT_VERSION")) = "CONST SMARTSTAT_VERSION" Then
      eqPos = InStr(1, line, "=", vbBinaryCompare)
      If eqPos > 0 Then
        ReadVersionConst = Trim(Mid(line, eqPos + 1))
      End If
      Exit Do
    End If
  Loop
  ts.Close
End Function

Sub EnsureParentFolder(ByVal filePath)
  Dim fso, parentPath
  Set fso = CreateObject("Scripting.FileSystemObject")
  parentPath = fso.GetParentFolderName(filePath)
  If Len(parentPath) = 0 Then Exit Sub
  Call EnsureFolderRecursive(parentPath)
End Sub

Sub EnsureFolderRecursive(ByVal folderPath)
  Dim fso, parentPath
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(folderPath) Then Exit Sub
  parentPath = fso.GetParentFolderName(folderPath)
  If Len(parentPath) > 0 Then
    If Not fso.FolderExists(parentPath) Then Call EnsureFolderRecursive(parentPath)
  End If
  On Error Resume Next
  fso.CreateFolder folderPath
  On Error GoTo 0
End Sub

Function GetNamedArg(ByVal argName, ByVal defaultValue)
  On Error Resume Next
  Dim namedArgs
  Set namedArgs = WScript.Arguments.Named
  If Err.Number = 0 Then
    If namedArgs.Exists(argName) Then
      GetNamedArg = CStr(namedArgs(argName))
      On Error GoTo 0
      Exit Function
    End If
  End If
  Err.Clear
  On Error GoTo 0
  GetNamedArg = CStr(defaultValue)
End Function
