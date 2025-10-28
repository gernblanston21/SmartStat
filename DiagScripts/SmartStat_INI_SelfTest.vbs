' SmartStat INI Self-Test v1.1 (ASCII-only)
Option Explicit
On Error Resume Next

Const BASE_DIR = "E:\EDRIVE\UNIVERSAL\SmartStat\"
Dim FILES: FILES = Array( _
  "SmartStat_Mappings.ini", _
  "SmartStat_Mappings.learn.ini", _
  "SmartStat_MappingsNBA.ini", _
  "SmartStat_MappingsNBA.learn.ini", _
  "SmartStat_MappingsNHL.ini", _
  "SmartStat_MappingsNHL.learn.ini", _
  "SmartStat_StaticOverrides.ini", _
  "SmartStat_TemplateConfig.ini" _
)

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Dim logPath: logPath = BASE_DIR & "DiagLogs\" & "SmartStat_INI_TestLog.txt"
Dim log: Set log = Nothing
Set log = fso.OpenTextFile(logPath, 2, True, 0) ' create/overwrite, ASCII
If Err.Number <> 0 Then
  Err.Clear
  Dim fb: fb = fso.GetParentFolderName(WScript.ScriptFullName) & "\SmartStat_INI_TestLog.txt"
  Set log = fso.OpenTextFile(fb, 2, True, 0)
  logPath = fb
End If

Sub W(s): log.WriteLine s: End Sub

W "=== SmartStat INI Self-Test ==="
W "Machine: " & CreateObject("WScript.Network").ComputerName
W "Date:    " & Now
W "Folder:  " & BASE_DIR
W ""

Dim i, name, path
Dim okCount: okCount = 0
Dim warnCount: warnCount = 0
Dim errCount: errCount = 0

For i = 0 To UBound(FILES)
  name = FILES(i)
  path = BASE_DIR & name

  W "------------------------------------------------------------"
  W "[" & name & "]"

  If Not fso.FileExists(path) Then
    W "ERROR: File not found at " & path
    errCount = errCount + 1
  Else
    Dim fileObj: Set fileObj = fso.GetFile(path)
    W "File found (" & fileObj.Size & " bytes)"

    Dim bomInfo: bomInfo = DetectBOM_ASCII(path)
    W "Encoding (BOM guess): " & bomInfo

    Dim text: text = ReadTextSmart_ASCII(path, bomInfo)
    If Err.Number <> 0 Then
      W "ERROR: Could not read file text. Err=" & Err.Description
      Err.Clear
      errCount = errCount + 1
    ElseIf Len(text) = 0 Then
      W "ERROR: File is empty or unreadable as text."
      errCount = errCount + 1
    Else
      ' Line-ending analysis
      Dim cCRLF, cLF, cCR, tmp
      cCRLF = CountOccurrences(text, vbCrLf)
      tmp = Replace(text, vbCrLf, "")
      cLF  = CountOccurrences(tmp, vbLf)
      tmp = Replace(text, vbCrLf, "")
      cCR  = CountOccurrences(tmp, vbCr)

      Dim eolNote
      If cCRLF > 0 And cLF = 0 And cCR = 0 Then
        eolNote = "Windows CRLF"
      ElseIf cCRLF = 0 And cLF > 0 Then
        eolNote = "LF only (Unix) detected"
        warnCount = warnCount + 1
      ElseIf cCRLF = 0 And cCR > 0 Then
        eolNote = "CR only (old Mac) detected"
        warnCount = warnCount + 1
      ElseIf cCRLF > 0 And cLF > 0 Then
        eolNote = "Mixed line endings (CRLF + LF)"
        warnCount = warnCount + 1
      Else
        eolNote = "No newline found or unknown"
        warnCount = warnCount + 1
      End If
      W "Line endings: " & eolNote

      ' Show first 5 logical lines
      Dim lines, j, maxShow
      lines = Split(NormalizeToLF(text), vbLf)
      maxShow = 5
      For j = 0 To UBound(lines)
        If j >= maxShow Then Exit For
        W "Line " & (j + 1) & ": " & SafeTrim(lines(j))
      Next

      ' Heuristic: check for INI-style sections
      If InStr(1, text, "[", vbTextCompare) = 0 Or InStr(1, text, "]", vbTextCompare) = 0 Then
        W "WARNING: No INI-like [sections] detected."
        warnCount = warnCount + 1
      End If

      okCount = okCount + 1
    End If
  End If

  W ""
Next

W "================ Summary ================"
W "OK files: " & okCount
W "Warnings: " & warnCount
W "Errors: " & errCount
W "Log saved to: " & logPath
log.Close

Dim msg
msg = "SmartStat INI Self-Test complete." & vbCrLf & _
      "OK: " & okCount & "   Warnings: " & warnCount & "   Errors: " & errCount & vbCrLf & _
      "Log: " & logPath
CreateObject("WScript.Shell").Popup msg, 8, "SmartStat INI Self-Test", 64
WScript.Quit 0

' -------- Helpers (ASCII only) --------
Function DetectBOM_ASCII(p)
  On Error Resume Next
  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Type = 1 ' binary
  stm.Open
  stm.LoadFromFile p
  Dim n: n = stm.Size
  Dim bytes
  If n >= 3 Then
    bytes = stm.Read(3) ' variant byte array
  ElseIf n = 2 Then
    bytes = stm.Read(2)
  ElseIf n = 1 Then
    bytes = stm.Read(1)
  Else
    bytes = Null
  End If
  stm.Close

  If IsArray(bytes) Then
    If UBound(bytes) >= 2 Then
      If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
        DetectBOM_ASCII = "UTF-8-BOM"
        Exit Function
      End If
    End If
    If UBound(bytes) >= 1 Then
      If bytes(0) = &HFF And bytes(1) = &HFE Then
        DetectBOM_ASCII = "UTF-16-LE"
        Exit Function
      ElseIf bytes(0) = &HFE And bytes(1) = &HFF Then
        DetectBOM_ASCII = "UTF-16-BE"
        Exit Function
      End If
    End If
  End If
  DetectBOM_ASCII = "No BOM (likely ANSI or UTF-8)"
End Function

Function ReadTextSmart_ASCII(p, bom)
  On Error Resume Next
  Dim cs
  If bom = "UTF-8-BOM" Then
    cs = "utf-8"
  ElseIf bom = "UTF-16-LE" Then
    cs = "utf-16"
  ElseIf bom = "UTF-16-BE" Then
    cs = "unicodeFFFE"
  Else
    cs = "utf-8" ' try UTF-8 no BOM first
  End If

  Dim stm: Set stm = CreateObject("ADODB.Stream")
  stm.Type = 2 ' text
  stm.Charset = cs
  stm.Open
  stm.LoadFromFile p
  Dim t: t = stm.ReadText
  stm.Close

  If Len(t) = 0 And bom = "No BOM (likely ANSI or UTF-8)" Then
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "windows-1252"
    stm.Open
    stm.LoadFromFile p
    t = stm.ReadText
    stm.Close
  End If
  ReadTextSmart_ASCII = t
End Function

Function CountOccurrences(ByVal haystack, ByVal needle)
  Dim pos, c: pos = 1: c = 0
  If Len(needle) = 0 Then CountOccurrences = 0: Exit Function
  Do
    pos = InStr(pos, haystack, needle, vbBinaryCompare)
    If pos = 0 Then Exit Do
    c = c + 1
    pos = pos + Len(needle)
  Loop
  CountOccurrences = c
End Function

Function NormalizeToLF(ByVal s)
  s = Replace(s, vbCrLf, vbLf)
  s = Replace(s, vbCr, vbLf)
  NormalizeToLF = s
End Function

Function SafeTrim(ByVal s)
  SafeTrim = Replace(Replace(s, vbTab, "    "), Chr(0), "")
End Function
