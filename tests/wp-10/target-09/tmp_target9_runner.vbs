Option Explicit

If WScript.Arguments.Count < 2 Then
  WScript.Echo "USAGE: cscript //nologo tmp_target9_runner.vbs <scriptPath> <scenario> [variant]"
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
  Case "B_AB"
    Call RunABDeterminism(variantId)
  Case "C_SINGLE"
    Call RunSingleWinner()
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

Function BuildLearn()
  Dim d : Set d = NewDictCI()
  d("fuzzy_threshold_short") = 0.6
  d("fuzzy_threshold") = 0.8
  d("ambiguous_score_delta") = 0.03
  d("phonetic_enable") = False
  d("allow_ambiguous_apply") = False
  d("deny_fuzzy") = Array("OBP", "OPS", "ERA", "WHIP", "WAR")
  Set BuildLearn = d
End Function

Sub BuildAmbiguousFixtures(ByVal variantId, ByRef catAlias, ByRef catMap, ByRef catPitchAlias, ByRef catPitchMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()
  Set catPitchAlias = NewDictCI()
  Set catPitchMap = NewDictCI()

  If UCase(CStr(variantId)) = "A" Then
    catAlias("abce") = "canon_one"
    catAlias("abcf") = "canon_two"
  Else
    catAlias("abcf") = "canon_two"
    catAlias("abce") = "canon_one"
  End If

  catMap("canon_one") = "FRAG_ONE"
  catMap("canon_two") = "FRAG_TWO"
End Sub

Sub BuildSingleWinnerFixtures(ByRef catAlias, ByRef catMap, ByRef catPitchAlias, ByRef catPitchMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()
  Set catPitchAlias = NewDictCI()
  Set catPitchMap = NewDictCI()

  catAlias("abce") = "canon_one"
  catAlias("zzzz") = "canon_two"

  catMap("canon_one") = "FRAG_ONE"
  catMap("canon_two") = "FRAG_TWO"
End Sub

Sub ResetState(ByVal mode)
  G_HARNESS_MODE = CStr(mode)
  Set CompilerContext = Nothing
  Set PlanValidationErrors = Nothing
  Set ApplyPlan = CreateObject("Scripting.Dictionary")
  Call EnsureAmbiguityContextEx(True, "tmp_target9_runner")
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

Sub RunStrictInitGuard()
  G_HARNESS_MODE = "HARNESS_STRICT"
  Set CompilerContext = Nothing
  Set PlanValidationErrors = Nothing

  Call EnsureAmbiguityContextEx(False, "tmp_target9_strict_guard")

  Dim hasKey : hasKey = HasAmbiguousContextInvalid()
  WScript.Echo "ASSERT_STRICT_NORMAL_INIT_NO_KEY=" & CStr(Not hasKey)
  WScript.Echo "STRICT_INIT_HAS_KEY=" & CStr(hasKey)
End Sub

Sub RunABDeterminism(ByVal variantId)
  Call ResetState("HARNESS_STRICT")

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap, catPitchAlias, catPitchMap
  Call BuildAmbiguousFixtures(variantId, catAlias, catMap, catPitchAlias, catPitchMap)

  Dim statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut
  statOut = "": conceptOut = "": isPitcher = False: usedHeuristic = False: acceptedBy = "": scoreOut = 0

  Dim ok
  ok = ResolveCategorySmart("abcd", False, learn, catAlias, catMap, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut)

  Call EmitResult(ok, statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut)
End Sub

Sub RunSingleWinner()
  Call ResetState("OFF")

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap, catPitchAlias, catPitchMap
  Call BuildSingleWinnerFixtures(catAlias, catMap, catPitchAlias, catPitchMap)

  Dim statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut
  statOut = "": conceptOut = "": isPitcher = False: usedHeuristic = False: acceptedBy = "": scoreOut = 0

  Dim ok
  ok = ResolveCategorySmart("abcd", False, learn, catAlias, catMap, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut)

  Call EmitResult(ok, statOut, conceptOut, isPitcher, usedHeuristic, acceptedBy, scoreOut)
End Sub
