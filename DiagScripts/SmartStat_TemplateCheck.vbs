' SmartStat_TemplateCheck.vbs
Option Explicit
On Error Resume Next

Const BASE = "E:\EDRIVE\UNIVERSAL\SmartStat\"
Dim tpl: tpl = InputBox("Enter the exact template name as it appears in Trio:", "SmartStat Template Check")
If Len(Trim(tpl)) = 0 Then WScript.Quit

Dim path: path = BASE & "SmartStat_TemplateConfig.ini"
Dim txt: txt = ReadText(path)
If Len(txt) = 0 Then
  MsgBox "Could not read " & path, vbExclamation, "SmartStat Template Check"
  WScript.Quit
End If

Dim block: block = ExtractBlock(txt, "TEMPLATE:" & tpl)
If Len(block) = 0 Then
  MsgBox "No [TEMPLATE:" & tpl & "] block found in SmartStat_TemplateConfig.ini", vbExclamation, "SmartStat Template Check"
  WScript.Quit
End If

Dim info: info = "Found [TEMPLATE:" & tpl & "]" & vbCrLf & vbCrLf & _
  "config_id=" & GetKey(block, "config_id") & vbCrLf & _
  "qualifier=" & GetKey(block, "qualifier") & vbCrLf & _
  "filter_tabfields=" & GetKey(block, "filter_tabfields") & vbCrLf & _
  "category_tabfields=" & GetKey(block, "category_tabfields") & vbCrLf & _
  "row_limit=" & GetKey(block, "row_limit") & vbCrLf & _
  "output_map=" & GetKey(block, "output_map")

MsgBox info, vbInformation, "SmartStat Template Check"

Function ReadText(p)
  Dim s: Set s = CreateObject("ADODB.Stream")
  s.Type = 2: s.Charset = "utf-8": s.Open
  s.LoadFromFile p
  ReadText = s.ReadText
  s.Close
End Function

Function ExtractBlock(ByVal allText, ByVal header)
  Dim u: u = UCase(allText)
  Dim startPos: startPos = InStr(1, u, "[" & UCase(header) & "]", vbTextCompare)
  If startPos = 0 Then Exit Function
  Dim endPos: endPos = InStr(startPos + 1, u, "[TEMPLATE:", vbTextCompare)
  If endPos = 0 Then endPos = Len(allText) + 1
  ExtractBlock = Trim(Mid(allText, startPos, endPos - startPos))
End Function

Function GetKey(ByVal blockText, ByVal keyName)
  Dim lines: lines = Split(Replace(Replace(blockText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
  Dim i, L, p
  For i = 0 To UBound(lines)
    L = Trim(lines(i))
    If L <> "" And Left(L,1) <> ";" And InStr(1, L, "=", vbTextCompare) > 0 Then
      p = InStr(1, L, "=", vbTextCompare)
      If UCase(Trim(Left(L, p-1))) = UCase(keyName) Then
        GetKey = Trim(Mid(L, p+1))
        Exit Function
      End If
    End If
  Next
  GetKey = "(missing)"
End Function
