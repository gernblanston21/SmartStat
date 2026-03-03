Option Explicit

If WScript.Arguments.Count < 2 Then
  WScript.Echo "USAGE: cscript //nologo tmp_target11_runner.vbs <scriptPath> <scenario> [variant]"
  WScript.Quit 2
End If

Dim scriptPath : scriptPath = CStr(WScript.Arguments(0))
Dim scenario  : scenario  = UCase(CStr(WScript.Arguments(1)))
Dim variantId : variantId = "A"
If WScript.Arguments.Count >= 3 Then variantId = UCase(CStr(WScript.Arguments(2)))

Call LoadSmartStat(scriptPath)

Select Case scenario
  Case "A_STRICT_INIT"
    Call RunStrictInitGuard()
  Case "B_AB_PRE"
    Call RunABPreBaseline()
  Case "B_AB_POST"
    Call RunABDeterminismPost(variantId)
  Case "C_SINGLE_PRE"
    Call RunSingleWinnerPre()
  Case "C_SINGLE_POST"
    Call RunSingleWinnerPost()
  Case Else
    WScript.Echo "UNKNOWN_SCENARIO=" & scenario
    WScript.Quit 3
End Select

Sub LoadSmartStat(ByVal path)
  Dim fso, ts, code
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(path) Then
    WScript.Echo "MISSING_SCRIPT=" & path
    WScript.Quit 4
  End If

  Set ts = fso.OpenTextFile(path, 1, False)
  code = ts.ReadAll
  ts.Close

  code = Replace(code, "Option Explicit" & vbCrLf, "")
  code = Replace(code, vbCrLf & "Call Main()" & vbCrLf, vbCrLf)
  code = Replace(code, "Call Main()" & vbCrLf, "")
  code = Replace(code, vbCrLf & "Call Main()", vbCrLf)

  ExecuteGlobal code
End Sub

Function NewDictCI()
  Dim d : Set d = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  d.CompareMode = 1
  On Error GoTo 0
  Set NewDictCI = d
End Function

Sub ResetState(ByVal mode)
  G_HARNESS_MODE = CStr(mode)
  Set CompilerContext = Nothing
  Set PlanValidationErrors = Nothing
  Set ApplyPlan = CreateObject("Scripting.Dictionary")
  Call EnsureAmbiguityContextEx(True, "tmp_target11_runner")
End Sub

Function HasAmbiguousContextInvalid()
  HasAmbiguousContextInvalid = False
  If (Not IsObject(PlanValidationErrors)) Then Exit Function
  If (PlanValidationErrors Is Nothing) Then Exit Function
  HasAmbiguousContextInvalid = PlanValidationErrors.Exists("AMBIGUOUS_CONTEXT_INVALID")
End Function

Function AmbiguityCount()
  AmbiguityCount = 0
  If (Not IsObject(CompilerContext)) Then Exit Function
  If (CompilerContext Is Nothing) Then Exit Function
  If Not CompilerContext.Exists("ambiguous") Then Exit Function
  If (Not IsObject(CompilerContext("ambiguous"))) Then Exit Function
  If UCase(CStr(TypeName(CompilerContext("ambiguous")))) <> "DICTIONARY" Then Exit Function
  AmbiguityCount = CompilerContext("ambiguous").Count
End Function

Function AmbiguitySummaryStable()
  AmbiguitySummaryStable = ""
  If AmbiguityCount() <= 0 Then Exit Function

  Dim d : Set d = CompilerContext("ambiguous")
  Dim keys, i, k
  keys = Transform_SortStringArrayTextBinary(d.Keys)
  If Not IsArray(keys) Then Exit Function

  For i = LBound(keys) To UBound(keys)
    k = CStr(keys(i))
    If Len(AmbiguitySummaryStable) > 0 Then AmbiguitySummaryStable = AmbiguitySummaryStable & " || "
    AmbiguitySummaryStable = AmbiguitySummaryStable & CStr(d(k))
  Next
End Function

Sub EmitResult(ByVal ok, ByVal statOut, ByVal conceptOut, ByVal isPitcher, ByVal usedHeuristic, ByVal acceptedBy, ByVal scoreOut)
  WScript.Echo "RET=" & CStr(ok)
  WScript.Echo "STAT_OUT=" & CStr(statOut)
  WScript.Echo "CONCEPT_OUT=" & CStr(conceptOut)
  WScript.Echo "IS_PITCHER=" & CStr(isPitcher)
  WScript.Echo "USED_HEUR=" & CStr(usedHeuristic)
  WScript.Echo "ACCEPTED_BY=" & CStr(acceptedBy)
  WScript.Echo "SCORE=" & ScoreStr(scoreOut)
  WScript.Echo "AMB_COUNT=" & CStr(AmbiguityCount())
  WScript.Echo "AMB_SUMMARY=" & CStr(AmbiguitySummaryStable())
End Sub

Sub EmitFromResult(ByVal r)
  Call EmitResult(CBool(r("ret")), CStr(r("stat")), CStr(r("concept")), CStr(r("is_pitcher")), CStr(r("used_heur")), CStr(r("accepted")), CDbl(r("score")))
End Sub

Function BuildIniCollision(ByVal variantId)
  Dim ini : Set ini = NewDictCI()
  Dim sec : Set sec = NewDictCI()

  If UCase(CStr(variantId)) = "A" Then
    sec("ALPHA BETA") = "VAL_SPACE"
    sec("ALPHA_BETA") = "VAL_UNDERSCORE"
  Else
    sec("ALPHA_BETA") = "VAL_UNDERSCORE"
    sec("ALPHA BETA") = "VAL_SPACE"
  End If

  Set ini("TEST_SEC") = sec
  Set BuildIniCollision = ini
End Function

Function BuildIniSingleWinner()
  Dim ini : Set ini = NewDictCI()
  Dim sec : Set sec = NewDictCI()

  sec("HITS") = "VAL_HITS"
  sec("WALKS") = "VAL_WALKS"

  Set ini("TEST_SEC") = sec
  Set BuildIniSingleWinner = ini
End Function

Sub LoadIniSectionDictNormalized_Pre(ini, sectionName, ByRef rawDict, ByRef normDict)
  Dim k, v
  Set rawDict = NewDictCI()
  Set normDict = NewDictCI()
  If ini Is Nothing Then Exit Sub
  If ini.Exists(sectionName) Then
    Dim sec: Set sec = ini(sectionName)
    For Each k In sec.Keys
      v = sec(k)
      If Not rawDict.Exists(k) Then rawDict.Add k, v
      Dim nk: nk = NormalizeKey(CStr(k))
      If Not normDict.Exists(nk) Then normDict.Add nk, v
    Next

    Dim sortedKeys: sortedKeys = Transform_SortStringArrayTextBinary(sec.Keys)
    If IsArray(sortedKeys) Then
      Dim strictMode: strictMode = Ambiguity_IsStrictHarness()
      Dim groupNormLookup, groupCount, groupKeys()
      Dim i, j, keyTxt, keyNormLookup
      Dim groupMatchCsv, groupWinner, groupWinnerNorm
      Dim firstVal, hasDiffVal

      groupNormLookup = ""
      groupCount = -1

      For i = LBound(sortedKeys) To UBound(sortedKeys)
        keyTxt = CStr(sortedKeys(i))
        keyNormLookup = NormalizeKeyForLookup(keyTxt)

        If groupCount >= 0 Then
          If StrComp(CStr(groupNormLookup), CStr(keyNormLookup), vbBinaryCompare) <> 0 Then
            If groupCount > 0 Then
              firstVal = CStr(sec(CStr(groupKeys(0))))
              hasDiffVal = False
              For j = 1 To groupCount
                If StrComp(firstVal, CStr(sec(CStr(groupKeys(j)))), vbBinaryCompare) <> 0 Then
                  hasDiffVal = True
                  Exit For
                End If
              Next

              If hasDiffVal Then
                groupMatchCsv = ""
                For j = 0 To groupCount
                  If Len(groupMatchCsv) > 0 Then groupMatchCsv = groupMatchCsv & "|"
                  groupMatchCsv = groupMatchCsv & CStr(groupKeys(j))
                Next

                If strictMode Then
                  Call Ambiguity_AddEx("ini", "normalized_collision", CStr(sectionName) & ":" & CStr(groupNormLookup), "top_tie=[" & groupMatchCsv & "]", "LoadIniSectionDictNormalized normalize collision fail-closed")
                Else
                  groupWinner = CStr(groupKeys(0))
                  groupWinnerNorm = NormalizeKey(CStr(groupWinner))
                  normDict(groupWinnerNorm) = sec(groupWinner)
                End If
              End If
            End If
            groupCount = -1
          End If
        End If

        If groupCount < 0 Then
          groupNormLookup = keyNormLookup
          groupCount = 0
          ReDim groupKeys(0)
          groupKeys(0) = keyTxt
        Else
          groupCount = groupCount + 1
          ReDim Preserve groupKeys(groupCount)
          groupKeys(groupCount) = keyTxt
        End If
      Next

      If groupCount > 0 Then
        firstVal = CStr(sec(CStr(groupKeys(0))))
        hasDiffVal = False
        For j = 1 To groupCount
          If StrComp(firstVal, CStr(sec(CStr(groupKeys(j)))), vbBinaryCompare) <> 0 Then
            hasDiffVal = True
            Exit For
          End If
        Next

        If hasDiffVal Then
          groupMatchCsv = ""
          For j = 0 To groupCount
            If Len(groupMatchCsv) > 0 Then groupMatchCsv = groupMatchCsv & "|"
            groupMatchCsv = groupMatchCsv & CStr(groupKeys(j))
          Next

          If strictMode Then
            Call Ambiguity_AddEx("ini", "normalized_collision", CStr(sectionName) & ":" & CStr(groupNormLookup), "top_tie=[" & groupMatchCsv & "]", "LoadIniSectionDictNormalized normalize collision fail-closed")
          Else
            groupWinner = CStr(groupKeys(0))
            groupWinnerNorm = NormalizeKey(CStr(groupWinner))
            normDict(groupWinnerNorm) = sec(groupWinner)
          End If
        End If
      End If
    End If
  End If
End Sub

Function EvalAB(ByVal usePostIngress, ByVal variantId)
  Call ResetState("HARNESS_STRICT")

  Dim ini : Set ini = BuildIniCollision(variantId)
  Dim rawDict, normDict
  If CBool(usePostIngress) Then
    Call LoadIniSectionDictNormalized(ini, "TEST_SEC", rawDict, normDict)
  Else
    Call LoadIniSectionDictNormalized_Pre(ini, "TEST_SEC", rawDict, normDict)
  End If

  Dim statOut : statOut = ""
  Dim wantedNormKey : wantedNormKey = NormalizeKey("ALPHA BETA")
  If IsObject(normDict) Then
    If normDict.Exists(wantedNormKey) Then statOut = CStr(normDict(wantedNormKey))
  End If

  Dim out : Set out = NewDictCI()
  out("ret") = CStr(AmbiguityCount() = 0)
  out("stat") = CStr(statOut)
  out("concept") = "INI.TEST_SEC"
  out("is_pitcher") = "False"
  out("used_heur") = "False"
  If AmbiguityCount() > 0 Then
    out("accepted") = "ambiguous/ini"
    out("score") = CDbl(0)
  Else
    out("accepted") = "direct"
    out("score") = CDbl(1)
  End If
  out("amb_count") = CStr(AmbiguityCount())
  out("amb_summary") = CStr(AmbiguitySummaryStable())
  Set EvalAB = out
End Function

Function EvalSingle(ByVal usePostIngress)
  Call ResetState("OFF")

  Dim ini : Set ini = BuildIniSingleWinner()
  Dim rawDict, normDict
  If CBool(usePostIngress) Then
    Call LoadIniSectionDictNormalized(ini, "TEST_SEC", rawDict, normDict)
  Else
    Call LoadIniSectionDictNormalized_Pre(ini, "TEST_SEC", rawDict, normDict)
  End If

  Dim statOut : statOut = ""
  If IsObject(normDict) Then
    If normDict.Exists("hits") Then statOut = CStr(normDict("hits"))
  End If

  Dim out : Set out = NewDictCI()
  out("ret") = CStr(Len(statOut) > 0 And AmbiguityCount() = 0)
  out("stat") = CStr(statOut)
  out("concept") = "INI.TEST_SEC"
  out("is_pitcher") = "False"
  out("used_heur") = "False"
  out("accepted") = "direct"
  out("score") = CDbl(1)
  out("amb_count") = CStr(AmbiguityCount())
  out("amb_summary") = CStr(AmbiguitySummaryStable())
  Set EvalSingle = out
End Function

Sub RunStrictInitGuard()
  G_HARNESS_MODE = "HARNESS_STRICT"
  Set CompilerContext = Nothing
  Set PlanValidationErrors = Nothing

  Call EnsureAmbiguityContextEx(False, "tmp_target11_strict_guard")

  Dim hasKey : hasKey = HasAmbiguousContextInvalid()

  Dim rAbPre, rAbPostA, rAbPostB, rSinglePre, rSinglePost
  Set rAbPre = EvalAB(False, "A")
  Set rAbPostA = EvalAB(True, "A")
  Set rAbPostB = EvalAB(True, "B")
  Set rSinglePre = EvalSingle(False)
  Set rSinglePost = EvalSingle(True)

  Dim parityAB, parityABPrePost, paritySingle, allPass
  parityAB = (CStr(rAbPostA("ret")) = CStr(rAbPostB("ret"))) And _
             (CStr(rAbPostA("accepted")) = CStr(rAbPostB("accepted"))) And _
             (CDbl(rAbPostA("score")) = CDbl(rAbPostB("score"))) And _
             (CStr(rAbPostA("amb_count")) = CStr(rAbPostB("amb_count"))) And _
             (CStr(rAbPostA("amb_summary")) = CStr(rAbPostB("amb_summary")))

  parityABPrePost = (CStr(rAbPre("ret")) = CStr(rAbPostA("ret"))) And _
                    (CStr(rAbPre("accepted")) = CStr(rAbPostA("accepted")))

  paritySingle = (CStr(rSinglePre("ret")) = CStr(rSinglePost("ret"))) And _
                 (CStr(rSinglePre("stat")) = CStr(rSinglePost("stat"))) And _
                 (CStr(rSinglePre("concept")) = CStr(rSinglePost("concept"))) And _
                 (CStr(rSinglePre("accepted")) = CStr(rSinglePost("accepted"))) And _
                 (CDbl(rSinglePre("score")) = CDbl(rSinglePost("score"))) And _
                 (CStr(rSinglePre("amb_count")) = CStr(rSinglePost("amb_count")))

  allPass = ((Not hasKey) And parityAB And parityABPrePost And paritySingle)

  WScript.Echo "ASSERT_STRICT_NORMAL_INIT_NO_KEY=" & CStr(allPass)
  WScript.Echo "STRICT_INIT_HAS_KEY=" & CStr(hasKey)
End Sub

Sub RunABPreBaseline()
  Dim r : Set r = EvalAB(False, "A")
  Call EmitFromResult(r)
End Sub

Sub RunABDeterminismPost(ByVal variantId)
  Dim r : Set r = EvalAB(True, variantId)
  Call EmitFromResult(r)
End Sub

Sub RunSingleWinnerPre()
  Dim r : Set r = EvalSingle(False)
  Call EmitFromResult(r)
End Sub

Sub RunSingleWinnerPost()
  Dim r : Set r = EvalSingle(True)
  Call EmitFromResult(r)
End Sub
