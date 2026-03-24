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

Public Function AdvisoryComposite_ResultLineExtended(ByRef resultObj)
    Dim hasAmbiguity, ambiguityFlagsText
    hasAmbiguity = "false"
    If resultObj.Exists("has_ambiguity") Then
        If CBool(resultObj("has_ambiguity")) Then hasAmbiguity = "true"
    End If
    ambiguityFlagsText = ""
    If resultObj.Exists("ambiguity_flags") Then
        ambiguityFlagsText = AC_JoinStringArray(resultObj("ambiguity_flags"), ";")
    End If

    AdvisoryComposite_ResultLineExtended = AdvisoryComposite_ResultLine(resultObj) & vbTab & _
                                           hasAmbiguity & vbTab & _
                                           ambiguityFlagsText
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

    Call AC_AttachAmbiguitySurface(baseResult, mapData, validations)

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
            baseResult("reason_code") = "DEFER_NON_APPROVED_PATTERN_WITH_VALID_MAPPING"
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

    Dim categoryTokens, categoryTargets, categoryAliases, qualifierTokens, qualifierTargets, qualifierAliases
    Set categoryTokens = mapData("CATEGORY_TOKEN")
    Set categoryTargets = mapData("CATEGORY_TARGET")
    Set categoryAliases = mapData("CATEGORY_ALIAS")
    Set qualifierTokens = mapData("QUALIFIER_TOKEN")
    Set qualifierTargets = mapData("QUALIFIER_TARGET")
    Set qualifierAliases = mapData("QUALIFIER_ALIAS")

    Dim rows(), i, token, norm, mappingType, canonicalToken, mappedTargetValue, isValid, canonicalNorm
    ReDim rows(UBound(components))

    For i = LBound(components) To UBound(components)
        token = CStr(components(i))
        norm = AC_NormalizeMappingToken(token)
        mappingType = "UNKNOWN"
        canonicalToken = ""
        mappedTargetValue = ""
        isValid = False

        If categoryTokens.Exists(norm) Then
            mappingType = "CATEGORY"
            canonicalToken = CStr(categoryTokens(norm))
            mappedTargetValue = CStr(categoryTargets(norm))
            isValid = True
        ElseIf categoryAliases.Exists(norm) Then
            canonicalNorm = CStr(categoryAliases(norm))
            If categoryTokens.Exists(canonicalNorm) Then
                mappingType = "CATEGORY"
                canonicalToken = CStr(categoryTokens(canonicalNorm))
                mappedTargetValue = CStr(categoryTargets(canonicalNorm))
                isValid = True
            End If
        ElseIf qualifierTokens.Exists(norm) Then
            mappingType = "QUALIFIER"
            canonicalToken = CStr(qualifierTokens(norm))
            mappedTargetValue = CStr(qualifierTargets(norm))
            isValid = True
        ElseIf qualifierAliases.Exists(norm) Then
            canonicalNorm = CStr(qualifierAliases(norm))
            If qualifierTokens.Exists(canonicalNorm) Then
                mappingType = "QUALIFIER"
                canonicalToken = CStr(qualifierTokens(canonicalNorm))
                mappedTargetValue = CStr(qualifierTargets(canonicalNorm))
                isValid = True
            End If
        End If

        Set rows(i) = AC_NewComponentValidation(token, norm, mappingType, canonicalToken, mappedTargetValue, isValid)
    Next

    AC_BuildComponentValidation = rows
End Function

Private Function AC_NewComponentValidation(ByVal token, ByVal normalizedToken, ByVal mappingType, ByVal canonicalToken, ByVal mappedTargetValue, ByVal isValid)
    Dim row
    Set row = CreateObject("Scripting.Dictionary")
    row.CompareMode = vbTextCompare
    row("token") = token
    row("normalized") = normalizedToken
    row("mapping_type") = mappingType
    row("canonical_token") = canonicalToken
    row("mapped_target_value") = mappedTargetValue
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

Private Sub AC_AttachAmbiguitySurface(ByRef baseResult, ByRef mapData, ByRef validations)
    Dim flagsSet
    Set flagsSet = CreateObject("Scripting.Dictionary")
    flagsSet.CompareMode = vbTextCompare

    Dim detailMap
    Set detailMap = CreateObject("Scripting.Dictionary")
    detailMap.CompareMode = vbTextCompare
    Dim detailIndex
    detailIndex = 0

    Dim components, i, token, norm, candidates
    components = baseResult("components")
    If IsArray(components) Then
        For i = LBound(components) To UBound(components)
            token = CStr(components(i))
            norm = AC_NormalizeMappingToken(token)
            candidates = AC_GetTokenCandidates(norm, mapData)
            Call AC_DetectTokenAmbiguity(token, candidates, flagsSet, detailMap, detailIndex)
        Next
    End If

    Call AC_DetectAliasCollision(validations, flagsSet, detailMap, detailIndex)
    Call AC_DetectSemanticDomainConflict(validations, flagsSet, detailMap, detailIndex)

    baseResult("ambiguity_flags") = AC_OrderedFlags(flagsSet)
    baseResult("ambiguity_details") = AC_DetailMapToArray(detailMap)
    baseResult("has_ambiguity") = (flagsSet.Count > 0)
End Sub

Private Sub AC_DetectTokenAmbiguity(ByVal token, ByRef candidates, ByRef flagsSet, ByRef detailMap, ByRef detailIndex)
    If Not AC_ArrayHasItems(candidates) Then Exit Sub

    Dim canonicalSet, targetSet, candidateLabels
    Set canonicalSet = CreateObject("Scripting.Dictionary")
    Set targetSet = CreateObject("Scripting.Dictionary")
    Set candidateLabels = CreateObject("Scripting.Dictionary")
    canonicalSet.CompareMode = vbTextCompare
    targetSet.CompareMode = vbTextCompare
    candidateLabels.CompareMode = vbTextCompare

    Dim hasCategory, hasQualifier, i, mappingType, canonicalToken, mappedTargetValue, label
    hasCategory = False
    hasQualifier = False

    For i = LBound(candidates) To UBound(candidates)
        mappingType = CStr(candidates(i)("mapping_type"))
        canonicalToken = CStr(candidates(i)("canonical_token"))
        mappedTargetValue = CStr(candidates(i)("mapped_target_value"))

        If mappingType = "CATEGORY" Then hasCategory = True
        If mappingType = "QUALIFIER" Then hasQualifier = True

        If Len(canonicalToken) > 0 Then
            If Not canonicalSet.Exists(canonicalToken) Then canonicalSet.Add canonicalToken, True
        End If
        If Len(mappedTargetValue) > 0 Then
            If Not targetSet.Exists(mappedTargetValue) Then targetSet.Add mappedTargetValue, True
        End If

        label = mappingType & ":" & canonicalToken & "=>" & mappedTargetValue
        If Not candidateLabels.Exists(label) Then candidateLabels.Add label, True
    Next

    If (canonicalSet.Count > 1) Or (targetSet.Count > 1) Then
        flagsSet("MULTI_MATCH") = True
        Call AC_AddAmbiguityDetail(detailMap, detailIndex, AC_NewDetail_TokenCandidates("MULTI_MATCH", token, AC_SortedKeys(candidateLabels)))
    End If

    If hasCategory And hasQualifier Then
        flagsSet("CROSS_DOMAIN_CONFLICT") = True
        Call AC_AddAmbiguityDetail(detailMap, detailIndex, AC_NewDetail_TokenCandidates("CROSS_DOMAIN_CONFLICT", token, AC_SortedKeys(candidateLabels)))
    End If
End Sub

Private Sub AC_DetectAliasCollision(ByRef validations, ByRef flagsSet, ByRef detailMap, ByRef detailIndex)
    If Not AC_ArrayHasItems(validations) Then Exit Sub

    Dim canonicalToTokens
    Set canonicalToTokens = CreateObject("Scripting.Dictionary")
    canonicalToTokens.CompareMode = vbTextCompare

    Dim i, canonicalToken, token, tokenSet
    For i = LBound(validations) To UBound(validations)
        If CBool(validations(i)("is_valid")) Then
            canonicalToken = CStr(validations(i)("canonical_token"))
            token = CStr(validations(i)("token"))
            If Len(canonicalToken) > 0 Then
                If Not canonicalToTokens.Exists(canonicalToken) Then
                    Set tokenSet = CreateObject("Scripting.Dictionary")
                    tokenSet.CompareMode = vbTextCompare
                    canonicalToTokens.Add canonicalToken, tokenSet
                End If
                Set tokenSet = canonicalToTokens(canonicalToken)
                If Not tokenSet.Exists(token) Then tokenSet.Add token, True
            End If
        End If
    Next

    Dim canonicalKeys, k
    canonicalKeys = AC_SortedKeys(canonicalToTokens)
    If AC_ArrayHasItems(canonicalKeys) Then
        For k = LBound(canonicalKeys) To UBound(canonicalKeys)
            canonicalToken = CStr(canonicalKeys(k))
            Set tokenSet = canonicalToTokens(canonicalToken)
            If tokenSet.Count > 1 Then
                flagsSet("ALIAS_COLLISION") = True
                Call AC_AddAmbiguityDetail(detailMap, detailIndex, AC_NewDetail_AliasCollision(canonicalToken, AC_SortedKeys(tokenSet)))
            End If
        Next
    End If
End Sub

Private Sub AC_DetectSemanticDomainConflict(ByRef validations, ByRef flagsSet, ByRef detailMap, ByRef detailIndex)
    If Not AC_ArrayHasItems(validations) Then Exit Sub

    Dim hitterTokens, pitcherTokens
    Set hitterTokens = CreateObject("Scripting.Dictionary")
    Set pitcherTokens = CreateObject("Scripting.Dictionary")
    hitterTokens.CompareMode = vbTextCompare
    pitcherTokens.CompareMode = vbTextCompare

    Dim i, mappedTargetValue, token
    For i = LBound(validations) To UBound(validations)
        If CBool(validations(i)("is_valid")) Then
            mappedTargetValue = CStr(validations(i)("mapped_target_value"))
            token = CStr(validations(i)("token"))
            If AC_IsExplicitPitcherMeasure(mappedTargetValue) Then
                If Not pitcherTokens.Exists(token) Then pitcherTokens.Add token, True
            ElseIf AC_IsExplicitHitterMeasure(mappedTargetValue) Then
                If Not hitterTokens.Exists(token) Then hitterTokens.Add token, True
            End If
        End If
    Next

    If (pitcherTokens.Count > 0) And (hitterTokens.Count > 0) Then
        flagsSet("SEMANTIC_DOMAIN_CONFLICT") = True
        Call AC_AddAmbiguityDetail(detailMap, detailIndex, AC_NewDetail_SemanticConflict(AC_SortedKeys(hitterTokens), AC_SortedKeys(pitcherTokens)))
    End If
End Sub

Private Function AC_IsExplicitPitcherMeasure(ByVal mappedTargetValue)
    AC_IsExplicitPitcherMeasure = (Left(LCase(CStr(mappedTargetValue)), 8) = "pitcher_")
End Function

Private Function AC_IsExplicitHitterMeasure(ByVal mappedTargetValue)
    AC_IsExplicitHitterMeasure = (Left(LCase(CStr(mappedTargetValue)), 8) = "batting_")
End Function

Private Function AC_GetTokenCandidates(ByVal normalizedToken, ByRef mapData)
    Dim categoryTokens, categoryTargets, categoryAliases, qualifierTokens, qualifierTargets, qualifierAliases
    Set categoryTokens = mapData("CATEGORY_TOKEN")
    Set categoryTargets = mapData("CATEGORY_TARGET")
    Set categoryAliases = mapData("CATEGORY_ALIAS")
    Set qualifierTokens = mapData("QUALIFIER_TOKEN")
    Set qualifierTargets = mapData("QUALIFIER_TARGET")
    Set qualifierAliases = mapData("QUALIFIER_ALIAS")

    Dim candidateMap
    Set candidateMap = CreateObject("Scripting.Dictionary")
    candidateMap.CompareMode = vbTextCompare

    Dim canonicalNorm
    If categoryTokens.Exists(normalizedToken) Then
        Call AC_AddCandidate(candidateMap, "CATEGORY", CStr(categoryTokens(normalizedToken)), CStr(categoryTargets(normalizedToken)))
    End If
    If categoryAliases.Exists(normalizedToken) Then
        canonicalNorm = CStr(categoryAliases(normalizedToken))
        If categoryTokens.Exists(canonicalNorm) Then
            Call AC_AddCandidate(candidateMap, "CATEGORY", CStr(categoryTokens(canonicalNorm)), CStr(categoryTargets(canonicalNorm)))
        End If
    End If
    If qualifierTokens.Exists(normalizedToken) Then
        Call AC_AddCandidate(candidateMap, "QUALIFIER", CStr(qualifierTokens(normalizedToken)), CStr(qualifierTargets(normalizedToken)))
    End If
    If qualifierAliases.Exists(normalizedToken) Then
        canonicalNorm = CStr(qualifierAliases(normalizedToken))
        If qualifierTokens.Exists(canonicalNorm) Then
            Call AC_AddCandidate(candidateMap, "QUALIFIER", CStr(qualifierTokens(canonicalNorm)), CStr(qualifierTargets(canonicalNorm)))
        End If
    End If

    AC_GetTokenCandidates = AC_CandidateMapToArray(candidateMap)
End Function

Private Sub AC_AddCandidate(ByRef candidateMap, ByVal mappingType, ByVal canonicalToken, ByVal mappedTargetValue)
    Dim key, row
    key = mappingType & "|" & canonicalToken & "|" & mappedTargetValue
    If candidateMap.Exists(key) Then Exit Sub

    Set row = CreateObject("Scripting.Dictionary")
    row.CompareMode = vbTextCompare
    row("mapping_type") = mappingType
    row("canonical_token") = canonicalToken
    row("mapped_target_value") = mappedTargetValue
    candidateMap.Add key, row
End Sub

Private Function AC_CandidateMapToArray(ByRef candidateMap)
    If candidateMap.Count = 0 Then
        AC_CandidateMapToArray = Array()
        Exit Function
    End If

    Dim keys, ordered, i, out(), idx
    keys = AC_SortedKeys(candidateMap)
    If Not AC_ArrayHasItems(keys) Then
        AC_CandidateMapToArray = Array()
        Exit Function
    End If

    ordered = keys
    ReDim out(UBound(ordered) - LBound(ordered))
    idx = 0
    For i = LBound(ordered) To UBound(ordered)
        Set out(idx) = candidateMap(CStr(ordered(i)))
        idx = idx + 1
    Next
    AC_CandidateMapToArray = out
End Function

Private Function AC_OrderedFlags(ByRef flagsSet)
    Dim ordered
    ordered = Array("MULTI_MATCH", "ALIAS_COLLISION", "CROSS_DOMAIN_CONFLICT", "SEMANTIC_DOMAIN_CONFLICT")

    Dim out(), idx, i, flagName
    idx = -1
    For i = LBound(ordered) To UBound(ordered)
        flagName = CStr(ordered(i))
        If flagsSet.Exists(flagName) Then
            idx = idx + 1
            ReDim Preserve out(idx)
            out(idx) = flagName
        End If
    Next

    If idx = -1 Then
        AC_OrderedFlags = Array()
    Else
        AC_OrderedFlags = out
    End If
End Function

Private Sub AC_AddAmbiguityDetail(ByRef detailMap, ByRef detailIndex, ByRef detailObj)
    detailMap.Add CStr(detailIndex), detailObj
    detailIndex = detailIndex + 1
End Sub

Private Function AC_DetailMapToArray(ByRef detailMap)
    If detailMap.Count = 0 Then
        AC_DetailMapToArray = Array()
        Exit Function
    End If

    Dim keys, i, out()
    keys = AC_SortedKeys(detailMap)
    ReDim out(UBound(keys) - LBound(keys))
    For i = LBound(keys) To UBound(keys)
        Set out(i - LBound(keys)) = detailMap(CStr(keys(i)))
    Next
    AC_DetailMapToArray = out
End Function

Private Function AC_NewDetail_TokenCandidates(ByVal detailType, ByVal token, ByRef candidates)
    Dim detail
    Set detail = CreateObject("Scripting.Dictionary")
    detail.CompareMode = vbTextCompare
    detail("type") = detailType
    detail("token") = token
    detail("candidates") = candidates
    Set AC_NewDetail_TokenCandidates = detail
End Function

Private Function AC_NewDetail_AliasCollision(ByVal canonicalToken, ByRef tokens)
    Dim detail
    Set detail = CreateObject("Scripting.Dictionary")
    detail.CompareMode = vbTextCompare
    detail("type") = "ALIAS_COLLISION"
    detail("canonical_token") = canonicalToken
    detail("tokens") = tokens
    Set AC_NewDetail_AliasCollision = detail
End Function

Private Function AC_NewDetail_SemanticConflict(ByRef hitterTokens, ByRef pitcherTokens)
    Dim detail
    Set detail = CreateObject("Scripting.Dictionary")
    detail.CompareMode = vbTextCompare
    detail("type") = "SEMANTIC_DOMAIN_CONFLICT"
    detail("hitter_tokens") = hitterTokens
    detail("pitcher_tokens") = pitcherTokens
    Set AC_NewDetail_SemanticConflict = detail
End Function

Private Function AC_SortedKeys(ByRef dictObj)
    If dictObj.Count = 0 Then
        AC_SortedKeys = Array()
        Exit Function
    End If

    Dim arr(), i, key
    ReDim arr(dictObj.Count - 1)
    i = 0
    For Each key In dictObj.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next
    Call AC_SortStringArray(arr)
    AC_SortedKeys = arr
End Function

Private Sub AC_SortStringArray(ByRef arr)
    Dim i, j, tmp
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(j)) < CStr(arr(i)) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next
    Next
End Sub

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

    Dim categoryTokens, categoryTargets, categoryAliases, qualifierTokens, qualifierTargets, qualifierAliases
    Set categoryTokens = CreateObject("Scripting.Dictionary")
    Set categoryTargets = CreateObject("Scripting.Dictionary")
    Set categoryAliases = CreateObject("Scripting.Dictionary")
    Set qualifierTokens = CreateObject("Scripting.Dictionary")
    Set qualifierTargets = CreateObject("Scripting.Dictionary")
    Set qualifierAliases = CreateObject("Scripting.Dictionary")
    categoryTokens.CompareMode = vbTextCompare
    categoryTargets.CompareMode = vbTextCompare
    categoryAliases.CompareMode = vbTextCompare
    qualifierTokens.CompareMode = vbTextCompare
    qualifierTargets.CompareMode = vbTextCompare
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
                        If Not categoryTokens.Exists(nk) Then
                            categoryTokens.Add nk, keyPart
                            categoryTargets.Add nk, valuePart
                        End If
                    Case "CATEGORY_TO_MEASURE_ALIASES", "CATEGORY_TO_MEASURE_PITCHER_ALIASES"
                        If Not categoryAliases.Exists(nk) Then categoryAliases.Add nk, nv
                    Case "QUALIFIER_TO_FILTER"
                        If Not qualifierTokens.Exists(nk) Then
                            qualifierTokens.Add nk, keyPart
                            qualifierTargets.Add nk, valuePart
                        End If
                    Case "QUALIFIER_TO_FILTER_ALIASES"
                        If Not qualifierAliases.Exists(nk) Then qualifierAliases.Add nk, nv
                End Select
            End If
        End If
    Loop
    stream.Close

    Call AC_AssertNoCategoryQualifierOverlap(categoryTokens, qualifierTokens)

    Dim mapData
    Set mapData = CreateObject("Scripting.Dictionary")
    mapData.CompareMode = vbTextCompare
    mapData.Add "CATEGORY_TOKEN", categoryTokens
    mapData.Add "CATEGORY_TARGET", categoryTargets
    mapData.Add "CATEGORY_ALIAS", categoryAliases
    mapData.Add "QUALIFIER_TOKEN", qualifierTokens
    mapData.Add "QUALIFIER_TARGET", qualifierTargets
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
            "CROSS_DOMAIN_CONFLICT flag=CROSS_DOMAIN_CONFLICT token=" & firstOverlap & _
            " :: Ambiguous mapping key exists in both CATEGORY and QUALIFIER domains"
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

Private Function AC_JoinStringArray(ByRef values, ByVal delimiter)
    Dim i, output
    output = ""
    If AC_ArrayHasItems(values) Then
        For i = LBound(values) To UBound(values)
            If i > LBound(values) Then output = output & delimiter
            output = output & CStr(values(i))
        Next
    End If
    AC_JoinStringArray = output
End Function

Private Function AC_ArrayHasItems(ByRef values)
    On Error Resume Next
    Dim lb, ub
    lb = LBound(values)
    ub = UBound(values)
    If Err.Number <> 0 Then
        AC_ArrayHasItems = False
    Else
        AC_ArrayHasItems = (ub >= lb)
    End If
    On Error GoTo 0
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
    If WScript.Arguments.Named.Exists("extended") Then
        WScript.Echo AdvisoryComposite_ResultLineExtended(cliResult)
    Else
        WScript.Echo AdvisoryComposite_ResultLine(cliResult)
    End If
End If
