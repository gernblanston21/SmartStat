Option Explicit

' SmartStat advisory-only composite recognizer.
' Scope is intentionally bounded for PASS_08:
' - strict recognition of PASS_07 approved PASS patterns
' - deterministic DEFER/REJECT fail-closed behavior
' - no runtime mutation, no apply/take/cue, no tabfield writes

Public Function AdvisoryComposite_Recognize(ByVal inputString)
    Dim normalized
    normalized = AC_NormalizeInput(inputString)

    Dim structureResult

    If Len(normalized) = 0 Then
        Set structureResult = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    ' Hard reject path-like structures before other checks.
    If InStr(normalized, ":/") > 0 Or InStr(normalized, ":\") > 0 Then
        Set structureResult = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    Dim approvedMeta
    approvedMeta = AC_GetApprovedMeta(normalized)
    If Len(approvedMeta) > 0 Then
        Set structureResult = AC_BuildApprovedPassResult(approvedMeta)
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    Dim hasSlash, hasComma
    hasSlash = (InStr(normalized, "/") > 0)
    hasComma = (InStr(normalized, ",") > 0)

    If hasSlash And hasComma Then
        Set structureResult = AC_NewResult("REJECT", "none", Array(), "REJECT_AMBIGUOUS_STRUCTURE")
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    If InStr(normalized, "//") > 0 Then
        Set structureResult = AC_NewResult("REJECT", "none", Array(), "REJECT_AMBIGUOUS_STRUCTURE")
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    If hasComma Then
        Set structureResult = AC_HandleCommaSequence(normalized)
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    If hasSlash Then
        Set structureResult = AC_HandleSlashStructure(normalized)
        Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
        Exit Function
    End If

    Set structureResult = AC_NewResult("REJECT", "none", Array(), "REJECT_UNSUPPORTED_PATTERN")
    Set AdvisoryComposite_Recognize = AC_ApplyMappingValidation(structureResult)
End Function

Public Function AdvisoryComposite_ResultLine(ByRef resultObj)
    Dim validCount, invalidCount
    validCount = 0
    invalidCount = 0
    If resultObj.Exists("valid_component_count") Then validCount = CInt(resultObj("valid_component_count"))
    If resultObj.Exists("invalid_component_count") Then invalidCount = CInt(resultObj("invalid_component_count"))

    AdvisoryComposite_ResultLine = CStr(resultObj("classification")) & vbTab & _
                                   CStr(resultObj("pattern_type")) & vbTab & _
                                   CStr(resultObj("reason_code")) & vbTab & _
                                   AC_JoinComponents(resultObj("components")) & vbTab & _
                                   CStr(validCount) & vbTab & _
                                   CStr(invalidCount)
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

Private Function AC_ApplyMappingValidation(ByRef baseResult)
    Dim mapData
    Set mapData = AC_LoadMappingData()

    Dim validations
    validations = AC_BuildComponentValidation(baseResult("components"), mapData)
    baseResult("component_validation") = validations

    Dim validCount, invalidCount
    validCount = AC_CountValid(validations)
    invalidCount = AC_CountInvalid(validations)
    baseResult("valid_component_count") = validCount
    baseResult("invalid_component_count") = invalidCount

    Dim baseClass
    baseClass = CStr(baseResult("classification"))

    If baseClass = "REJECT" Then
        Set AC_ApplyMappingValidation = baseResult
        Exit Function
    End If

    If validCount = 0 Then
        baseResult("classification") = "REJECT"
        baseResult("reason_code") = "REJECT_NO_VALID_COMPONENTS"
        Set AC_ApplyMappingValidation = baseResult
        Exit Function
    End If

    If invalidCount = 0 Then
        If baseClass = "PASS" Then
            baseResult("classification") = "PASS"
            baseResult("reason_code") = "PASS_VALIDATED_MAPPING"
        Else
            ' Preserve PASS_08 bounded structure authority for non-approved shapes.
            baseResult("classification") = "DEFER"
            baseResult("reason_code") = "DEFER_PARTIAL_MAPPING"
        End If
        Set AC_ApplyMappingValidation = baseResult
        Exit Function
    End If

    baseResult("classification") = "DEFER"
    baseResult("reason_code") = "DEFER_PARTIAL_MAPPING"
    Set AC_ApplyMappingValidation = baseResult
End Function

Private Function AC_BuildComponentValidation(ByRef components, ByRef mapData)
    If Not IsArray(components) Then
        AC_BuildComponentValidation = Array()
        Exit Function
    End If

    Dim categories, categoryAliases, qualifiers, qualifierAliases
    Set categories = mapData("CATEGORY")
    Set categoryAliases = mapData("CATEGORY_ALIAS")
    Set qualifiers = mapData("QUALIFIER")
    Set qualifierAliases = mapData("QUALIFIER_ALIAS")

    Dim rows(), i, token, norm, mappingType, mappingKey, isValid, canonicalNorm
    ReDim rows(UBound(components))

    For i = LBound(components) To UBound(components)
        token = CStr(components(i))
        norm = AC_NormalizeMappingToken(token)
        mappingType = "UNKNOWN"
        mappingKey = ""
        isValid = False

        If categories.Exists(norm) Then
            mappingType = "CATEGORY"
            mappingKey = CStr(categories(norm))
            isValid = True
        ElseIf categoryAliases.Exists(norm) Then
            canonicalNorm = CStr(categoryAliases(norm))
            If categories.Exists(canonicalNorm) Then
                mappingType = "CATEGORY"
                mappingKey = CStr(categories(canonicalNorm))
                isValid = True
            End If
        ElseIf qualifiers.Exists(norm) Then
            mappingType = "QUALIFIER"
            mappingKey = CStr(qualifiers(norm))
            isValid = True
        ElseIf qualifierAliases.Exists(norm) Then
            canonicalNorm = CStr(qualifierAliases(norm))
            If qualifiers.Exists(canonicalNorm) Then
                mappingType = "QUALIFIER"
                mappingKey = CStr(qualifiers(canonicalNorm))
                isValid = True
            End If
        End If

        Set rows(i) = AC_NewComponentValidation(token, norm, mappingType, mappingKey, isValid)
    Next

    AC_BuildComponentValidation = rows
End Function

Private Function AC_NewComponentValidation(ByVal token, ByVal normalizedToken, ByVal mappingType, ByVal mappingKey, ByVal isValid)
    Dim row
    Set row = CreateObject("Scripting.Dictionary")
    row.CompareMode = vbTextCompare
    row("token") = token
    row("normalized") = normalizedToken
    row("mapping_type") = mappingType
    row("mapping_key") = mappingKey
    row("is_valid") = CBool(isValid)
    Set AC_NewComponentValidation = row
End Function

Private Function AC_CountValid(ByRef validations)
    Dim i, n
    n = 0
    If IsArray(validations) Then
        For i = LBound(validations) To UBound(validations)
            If CBool(validations(i)("is_valid")) Then n = n + 1
        Next
    End If
    AC_CountValid = n
End Function

Private Function AC_CountInvalid(ByRef validations)
    Dim i, n
    n = 0
    If IsArray(validations) Then
        For i = LBound(validations) To UBound(validations)
            If Not CBool(validations(i)("is_valid")) Then n = n + 1
        Next
    End If
    AC_CountInvalid = n
End Function

Private Function AC_LoadMappingData()
    Dim fso, mappingPath, stream
    Set fso = CreateObject("Scripting.FileSystemObject")
    mappingPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\SmartStat_Mappings.ini"

    If Not fso.FileExists(mappingPath) Then
        Err.Raise vbObjectError + 8100, "AC_LoadMappingData", "SmartStat_Mappings.ini is missing: " & mappingPath
    End If

    On Error Resume Next
    Set stream = fso.OpenTextFile(mappingPath, 1, False)
    If Err.Number <> 0 Then
        Dim openErr
        openErr = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 8101, "AC_LoadMappingData", "Unable to read SmartStat_Mappings.ini: " & openErr
    End If
    On Error GoTo 0

    Dim categories, categoryAliases, qualifiers, qualifierAliases
    Set categories = CreateObject("Scripting.Dictionary")
    Set categoryAliases = CreateObject("Scripting.Dictionary")
    Set qualifiers = CreateObject("Scripting.Dictionary")
    Set qualifierAliases = CreateObject("Scripting.Dictionary")
    categories.CompareMode = vbTextCompare
    categoryAliases.CompareMode = vbTextCompare
    qualifiers.CompareMode = vbTextCompare
    qualifierAliases.CompareMode = vbTextCompare

    Dim section, line, eqPos, keyPart, valuePart, nk, nv
    section = ""
    Do While Not stream.AtEndOfStream
        line = Trim(stream.ReadLine)
        If Len(line) = 0 Then
            ' continue
        ElseIf Left(line, 1) = ";" Then
            ' continue
        ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
            section = UCase(Mid(line, 2, Len(line) - 2))
        Else
            eqPos = InStr(line, "=")
            If eqPos > 0 Then
                keyPart = Trim(Left(line, eqPos - 1))
                valuePart = Trim(Mid(line, eqPos + 1))
                nk = AC_NormalizeMappingToken(keyPart)
                nv = AC_NormalizeMappingToken(valuePart)

                Select Case section
                    Case "CATEGORY_TO_MEASURE", "CATEGORY_TO_MEASURE_PITCHER"
                        If Not categories.Exists(nk) Then categories.Add nk, keyPart
                    Case "CATEGORY_TO_MEASURE_ALIASES", "CATEGORY_TO_MEASURE_PITCHER_ALIASES"
                        If Not categoryAliases.Exists(nk) Then categoryAliases.Add nk, nv
                    Case "QUALIFIER_TO_FILTER"
                        If Not qualifiers.Exists(nk) Then qualifiers.Add nk, keyPart
                    Case "QUALIFIER_TO_FILTER_ALIASES"
                        If Not qualifierAliases.Exists(nk) Then qualifierAliases.Add nk, nv
                End Select
            End If
        End If
    Loop
    stream.Close

    Call AC_AssertNoCategoryQualifierOverlap(categories, qualifiers)

    Dim mapData
    Set mapData = CreateObject("Scripting.Dictionary")
    mapData.CompareMode = vbTextCompare
    mapData.Add "CATEGORY", categories
    mapData.Add "CATEGORY_ALIAS", categoryAliases
    mapData.Add "QUALIFIER", qualifiers
    mapData.Add "QUALIFIER_ALIAS", qualifierAliases
    Set AC_LoadMappingData = mapData
End Function

Private Sub AC_AssertNoCategoryQualifierOverlap(ByRef categories, ByRef qualifiers)
    Dim key, firstOverlap
    firstOverlap = ""
    For Each key In categories.Keys
        If qualifiers.Exists(CStr(key)) Then
            firstOverlap = CStr(key)
            Exit For
        End If
    Next

    If Len(firstOverlap) > 0 Then
        Err.Raise vbObjectError + 8102, "AC_AssertNoCategoryQualifierOverlap", _
            "Ambiguous mapping key exists in both CATEGORY and QUALIFIER domains: " & firstOverlap
    End If
End Sub

Private Function AC_NormalizeMappingToken(ByVal token)
    AC_NormalizeMappingToken = Replace(AC_NormalizeInput(token), " ", "_")
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
