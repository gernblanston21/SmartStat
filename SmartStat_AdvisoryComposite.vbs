Option Explicit

' SmartStat advisory-only composite recognizer.
' Scope is intentionally bounded for PASS_08:
' - strict recognition of PASS_07 approved PASS patterns
' - deterministic DEFER/REJECT fail-closed behavior
' - no runtime mutation, no apply/take/cue, no tabfield writes

Public Function AdvisoryComposite_Recognize(ByVal inputString)
    Dim normalized
    normalized = AC_NormalizeInput(inputString)

    If Len(normalized) = 0 Then
        Set AdvisoryComposite_Recognize = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Exit Function
    End If

    ' Hard reject path-like structures before other checks.
    If InStr(normalized, ":/") > 0 Or InStr(normalized, ":\") > 0 Then
        Set AdvisoryComposite_Recognize = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Exit Function
    End If

    Dim approvedMeta
    approvedMeta = AC_GetApprovedMeta(normalized)
    If Len(approvedMeta) > 0 Then
        Set AdvisoryComposite_Recognize = AC_BuildApprovedPassResult(approvedMeta)
        Exit Function
    End If

    Dim hasSlash, hasComma
    hasSlash = (InStr(normalized, "/") > 0)
    hasComma = (InStr(normalized, ",") > 0)

    If hasSlash And hasComma Then
        Set AdvisoryComposite_Recognize = AC_NewResult("REJECT", "none", Array(), "REJECT_AMBIGUOUS_STRUCTURE")
        Exit Function
    End If

    If InStr(normalized, "//") > 0 Then
        Set AdvisoryComposite_Recognize = AC_NewResult("REJECT", "none", Array(), "REJECT_AMBIGUOUS_STRUCTURE")
        Exit Function
    End If

    If hasComma Then
        Set AdvisoryComposite_Recognize = AC_HandleCommaSequence(normalized)
        Exit Function
    End If

    If hasSlash Then
        Set AdvisoryComposite_Recognize = AC_HandleSlashStructure(normalized)
        Exit Function
    End If

    Set AdvisoryComposite_Recognize = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
End Function

Public Function AdvisoryComposite_ResultLine(ByRef resultObj)
    AdvisoryComposite_ResultLine = CStr(resultObj("classification")) & vbTab & _
                                   CStr(resultObj("pattern_type")) & vbTab & _
                                   CStr(resultObj("reason_code")) & vbTab & _
                                   AC_JoinComponents(resultObj("components"))
End Function

Private Function AC_HandleCommaSequence(ByVal normalized)
    Dim rawParts, parts
    rawParts = Split(normalized, ",")
    parts = AC_SanitizeParts(rawParts)

    If Not IsArray(parts) Then
        Set AC_HandleCommaSequence = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Exit Function
    End If

    Dim count
    count = UBound(parts) - LBound(parts) + 1
    If count >= 2 And count <= 3 Then
        Set AC_HandleCommaSequence = AC_NewResult("DEFER", "multi_measure_sequence", parts, "DEFER_PARTIAL_MATCH")
    Else
        Set AC_HandleCommaSequence = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
    End If
End Function

Private Function AC_HandleSlashStructure(ByVal normalized)
    If Left(normalized, 1) = "/" Or Right(normalized, 1) = "/" Then
        Set AC_HandleSlashStructure = AC_NewResult("REJECT", "none", Array(), "REJECT_AMBIGUOUS_STRUCTURE")
        Exit Function
    End If

    Dim parentheticalParts
    parentheticalParts = AC_ParseSlashParenthetical(normalized)
    If IsArray(parentheticalParts) Then
        Set AC_HandleSlashStructure = AC_NewResult("DEFER", "multi_measure_sequence", parentheticalParts, "DEFER_PARTIAL_MATCH")
        Exit Function
    End If

    Dim rawParts, parts, count
    rawParts = Split(normalized, "/")
    parts = AC_SanitizeParts(rawParts)

    If Not IsArray(parts) Then
        Set AC_HandleSlashStructure = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Exit Function
    End If

    count = UBound(parts) - LBound(parts) + 1
    If count = 2 Then
        Set AC_HandleSlashStructure = AC_NewResult("DEFER", "ratio_or_slash", parts, "DEFER_PARTIAL_MATCH")
    ElseIf count = 3 Then
        Set AC_HandleSlashStructure = AC_NewResult("DEFER", "multi_measure_sequence", parts, "DEFER_PARTIAL_MATCH")
    Else
        Set AC_HandleSlashStructure = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
    End If
End Function

Private Function AC_BuildApprovedPassResult(ByVal approvedMeta)
    Dim metaParts, patternType, componentText, components
    metaParts = Split(approvedMeta, "|")
    patternType = CStr(metaParts(0))
    componentText = CStr(metaParts(1))
    components = Split(componentText, ";")
    Set AC_BuildApprovedPassResult = AC_NewResult("PASS", patternType, components, "PASS_APPROVED_PATTERN")
End Function

Private Function AC_GetApprovedMeta(ByVal normalized)
    Dim approved
    Set approved = CreateObject("Scripting.Dictionary")
    approved.CompareMode = vbTextCompare

    ' Frozen PASS_07 approved implementation set.
    approved("K/BB") = "ratio_or_slash|K;BB"
    approved("BB/K") = "ratio_or_slash|BB;K"
    approved("STRIKEOUTS/BB") = "ratio_or_slash|STRIKEOUTS;BB"
    approved("AVG/HR/RBI") = "multi_measure_sequence|AVG;HR;RBI"
    approved("OBP/SLG/OPS") = "multi_measure_sequence|OBP;SLG;OPS"
    approved("PA,HR,RBI") = "multi_measure_sequence|PA;HR;RBI"
    approved("HR/PA (OBP)") = "multi_measure_sequence|HR;PA;OBP"

    If approved.Exists(normalized) Then
        AC_GetApprovedMeta = CStr(approved(normalized))
    Else
        AC_GetApprovedMeta = ""
    End If
End Function

Private Function AC_ParseSlashParenthetical(ByVal normalized)
    Dim re, matches
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = "^\s*(.+?)\s*/\s*(.+?)\s*\(\s*(.+?)\s*\)\s*$"

    If re.Test(normalized) Then
        Set matches = re.Execute(normalized)
        Dim comps(2)
        comps(0) = AC_NormalizeInput(matches(0).SubMatches(0))
        comps(1) = AC_NormalizeInput(matches(0).SubMatches(1))
        comps(2) = AC_NormalizeInput(matches(0).SubMatches(2))
        If Len(comps(0)) > 0 And Len(comps(1)) > 0 And Len(comps(2)) > 0 Then
            AC_ParseSlashParenthetical = comps
            Exit Function
        End If
    End If

    AC_ParseSlashParenthetical = Empty
End Function

Private Function AC_SanitizeParts(ByRef rawParts)
    Dim i, item, clean(), idx
    ReDim clean(UBound(rawParts))
    idx = -1

    For i = 0 To UBound(rawParts)
        item = AC_NormalizeInput(rawParts(i))
        If Len(item) = 0 Then
            AC_SanitizeParts = Empty
            Exit Function
        End If
        idx = idx + 1
        clean(idx) = item
    Next

    ReDim Preserve clean(idx)
    AC_SanitizeParts = clean
End Function

Private Function AC_NormalizeInput(ByVal value)
    AC_NormalizeInput = UCase(Trim(CStr(value)))
End Function

Private Function AC_JoinComponents(ByRef components)
    Dim i, output
    output = ""
    If IsArray(components) Then
        For i = LBound(components) To UBound(components)
            If i > LBound(components) Then
                output = output & ";"
            End If
            output = output & CStr(components(i))
        Next
    End If
    AC_JoinComponents = output
End Function

Private Function AC_NewResult(ByVal classification, ByVal patternType, ByRef components, ByVal reasonCode)
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("classification") = classification
    result("pattern_type") = patternType
    result("components") = components
    result("reason_code") = reasonCode
    Set AC_NewResult = result
End Function

' Optional CLI mode for harness use:
' cscript //nologo SmartStat_AdvisoryComposite.vbs /input:"K/BB"
If WScript.Arguments.Named.Exists("input") Then
    Dim cliInput, cliResult
    cliInput = WScript.Arguments.Named("input")
    Set cliResult = AdvisoryComposite_Recognize(cliInput)
    WScript.Echo AdvisoryComposite_ResultLine(cliResult)
End If
