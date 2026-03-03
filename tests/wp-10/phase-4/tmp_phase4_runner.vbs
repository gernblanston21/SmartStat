Option Explicit

If WScript.Arguments.Count < 3 Then
  WScript.Echo "USAGE: cscript //nologo tmp_phase4_runner.vbs <scriptPath> <scenario> <mode>"
  WScript.Quit 2
End If

Dim scriptPath : scriptPath = CStr(WScript.Arguments(0))
Dim scenario : scenario = UCase(CStr(WScript.Arguments(1)))
Dim modeIn : modeIn = UCase(CStr(WScript.Arguments(2)))
Dim modeNorm : modeNorm = NormalizeMode(modeIn)

If Len(modeNorm) = 0 Then
  WScript.Echo "INVALID_MODE=" & modeIn
  WScript.Quit 5
End If

Call LoadSmartStat(scriptPath)

Dim r
Set r = Nothing

Select Case scenario
  Case "SUCCESS"
    Set r = EvalSingleWinner(modeNorm)
  Case "AMBIGUITY"
    Set r = EvalAmbiguousTie(modeNorm)
  Case Else
    WScript.Echo "UNKNOWN_SCENARIO=" & scenario
    WScript.Quit 3
End Select

Call EmitFromResult(r)

Function NormalizeMode(ByVal modeTxt)
  Dim t : t = UCase(Trim(CStr(modeTxt)))
  Select Case t
    Case "HARNESS_STRICT"
      NormalizeMode = "HARNESS_STRICT"
    Case "HARNESS", "OFF"
      ' OFF is accepted as alias for non-strict harness behavior.
      NormalizeMode = "HARNESS"
    Case Else
      NormalizeMode = ""
  End Select
End Function

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
  Call EnsureAmbiguityContextEx(True, "tmp_phase4_runner")
End Sub

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

Sub BuildSingleWinnerFixtures(ByRef catAlias, ByRef catMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()

  catAlias("abce") = "canon_one"
  catAlias("zzzz") = "canon_two"

  catMap("canon_one") = "FRAG_ONE"
  catMap("canon_two") = "FRAG_TWO"
End Sub

Sub BuildAmbiguousFixtures(ByRef catAlias, ByRef catMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()

  catAlias("abce") = "canon_one"
  catAlias("abcf") = "canon_two"

  catMap("canon_one") = "FRAG_ONE"
  catMap("canon_two") = "FRAG_TWO"
End Sub

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

Function EvalSingleWinner(ByVal modeTxt)
  Call ResetState(modeTxt)

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap
  Call BuildSingleWinnerFixtures(catAlias, catMap)

  Dim bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv, ok
  bestKey = "": bestScore = 0
  altKey = "": altScore = 0
  isAmb = False: hasTopTie = False: tieCsv = ""

  ok = HeuristicPickWithAlt("abcd", catAlias, learn, bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv)

  Dim ret, statOut, conceptOut, acceptedBy
  ret = CBool(ok)
  statOut = ""
  conceptOut = ""
  acceptedBy = "fuzzy/alias"

  If CBool(isAmb) Then
    Call Ambiguity_AddEx("category", "alias", "abcd", "top_tie=[" & tieCsv & "]", "HeuristicPickWithAlt top-score tie fail-closed")
    ret = False
    acceptedBy = "ambiguous/alias"
  ElseIf ret Then
    Dim canon
    canon = CStr(catAlias(bestKey))
    If catMap.Exists(canon) Then
      statOut = CStr(catMap(canon))
      conceptOut = "CATEGORY." & canon
    End If
  End If

  Dim out : Set out = NewDictCI()
  out("ret") = CStr(ret)
  out("stat") = CStr(statOut)
  out("concept") = CStr(conceptOut)
  out("is_pitcher") = "False"
  out("used_heur") = "True"
  out("accepted") = CStr(acceptedBy)
  out("score") = CDbl(bestScore)
  out("amb_count") = CStr(AmbiguityCount())
  out("amb_summary") = CStr(AmbiguitySummaryStable())
  Set EvalSingleWinner = out
End Function

Function EvalAmbiguousTie(ByVal modeTxt)
  Call ResetState(modeTxt)

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap
  Call BuildAmbiguousFixtures(catAlias, catMap)

  Dim bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv, ok
  bestKey = "": bestScore = 0
  altKey = "": altScore = 0
  isAmb = False: hasTopTie = False: tieCsv = ""

  ok = HeuristicPickWithAlt("abcd", catAlias, learn, bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv)

  Dim ret, statOut, conceptOut, acceptedBy
  ret = CBool(ok)
  statOut = ""
  conceptOut = ""
  acceptedBy = "fuzzy/alias"

  If CBool(isAmb) Then
    Call Ambiguity_AddEx("category", "alias", "abcd", "top_tie=[" & tieCsv & "]", "HeuristicPickWithAlt top-score tie fail-closed")
    ret = False
    acceptedBy = "ambiguous/alias"
  ElseIf ret Then
    Dim canon
    canon = CStr(catAlias(bestKey))
    If catMap.Exists(canon) Then
      statOut = CStr(catMap(canon))
      conceptOut = "CATEGORY." & canon
    End If
  End If

  Dim out : Set out = NewDictCI()
  out("ret") = CStr(ret)
  out("stat") = CStr(statOut)
  out("concept") = CStr(conceptOut)
  out("is_pitcher") = "False"
  out("used_heur") = "True"
  out("accepted") = CStr(acceptedBy)
  out("score") = CDbl(bestScore)
  out("amb_count") = CStr(AmbiguityCount())
  out("amb_summary") = CStr(AmbiguitySummaryStable())
  Set EvalAmbiguousTie = out
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
