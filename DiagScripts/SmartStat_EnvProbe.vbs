' SmartStat_EnvProbe.vbs
Option Explicit
On Error Resume Next

Const BASE = "E:\EDRIVE\UNIVERSAL\SmartStat\"
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim log: Set log = fso.OpenTextFile(BASE & "DiagLogs\" & "SmartStat_EnvProbe_Log.txt", 2, True, 0)

Sub W(s): log.WriteLine s: End Sub

W "=== SmartStat Env Probe ==="
W "Date: " & Now
W "Base: " & BASE
W ""

' 1) ADODB.Stream check
Dim okADO: okADO = False
Dim stm
On Error Resume Next
Set stm = CreateObject("ADODB.Stream")
If Err.Number = 0 Then
  okADO = True
  stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.WriteText "test": stm.Close
  W "ADODB.Stream: OK"
Else
  W "ADODB.Stream: MISSING (" & Err.Description & ")"
End If
Err.Clear

' 2) INI presence & readable
Dim files: files = Array( _
  "SmartStat_Mappings.ini", _
  "SmartStat_Mappings.learn.ini", _
  "SmartStat_MappingsNBA.ini", _
  "SmartStat_MappingsNBA.learn.ini", _
  "SmartStat_MappingsNHL.ini", _
  "SmartStat_MappingsNHL.learn.ini", _
  "SmartStat_StaticOverrides.ini", _
  "SmartStat_TemplateConfig.ini" _
)

Dim i, p, txt
For i = 0 To UBound(files)
  p = BASE & files(i)
  If Not fso.FileExists(p) Then
    W "[" & files(i) & "] MISSING: " & p
  Else
    txt = ReadText(p)
    If Len(txt) = 0 Then
      W "[" & files(i) & "] READ FAIL (empty or blocked)"
    Else
      W "[" & files(i) & "] OK (" & fso.GetFile(p).Size & " bytes)"
    End If
  End If
Next

' 3) Template count + sample names
Dim cfg: cfg = ReadText(BASE & "SmartStat_TemplateConfig.ini")
If Len(cfg) > 0 Then
  Dim lines: lines = Split(Replace(Replace(cfg, vbCrLf, vbLf), vbCr, vbLf), vbLf)
  Dim count: count = 0
  Dim sample: sample: sample = ""
  For i = 0 To UBound(lines)
    If Left(UCase(Trim(lines(i))), 10) = "[TEMPLATE:" Then
      count = count + 1
      If Len(sample) < 200 Then sample = sample & Trim(lines(i)) & "; "
    End If
  Next
  W ""
  W "Template blocks found: " & count
  If count > 0 Then W "Sample: " & sample
Else
  W ""
  W "TemplateConfig read failed."
End If

W ""
W "Probe complete."
log.Close

CreateObject("WScript.Shell").Popup "Env probe done. See SmartStat_EnvProbe_Log.txt", 6, "SmartStat Env", 64

Function ReadText(p)
  On Error Resume Next
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2: s.Charset = "utf-8"
  s.Open: s.LoadFromFile p
  ReadText = s.ReadText
  s.Close
End Function
