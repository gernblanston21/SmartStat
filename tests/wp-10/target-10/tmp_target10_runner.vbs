Option Explicit

If WScript.Arguments.Count < 2 Then
  WScript.Echo "USAGE: cscript //nologo tmp_target10_runner.vbs <scriptPath> <scenario> [variant]"
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

Sub BuildAmbiguousFixtures(ByVal variantId, ByRef catAlias, ByRef catMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()

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

Sub BuildSingleWinnerFixtures(ByRef catAlias, ByRef catMap)
  Set catAlias = NewDictCI()
  Set catMap = NewDictCI()

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
  Call EnsureAmbiguityContextEx(True, "tmp_target10_runner")
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

Function DuplicatePreservationPass()
  DuplicatePreservationPass = False

  Dim list : Set list = CreateObject("System.Collections.ArrayList")
  list.Add "hits"
  list.Add "hits"
  list.Add "hats"

  Dim arr, i, nOut, hitsCount
  arr = Fuzzy_NormalizeCandidateKeysForScan(list)

  nOut = 0
  hitsCount = 0
  If IsArray(arr) Then
    nOut = UBound(arr) - LBound(arr) + 1
    For i = LBound(arr) To UBound(arr)
      If LCase(CStr(arr(i))) = "hits" Then hitsCount = hitsCount + 1
    Next
  End If

  DuplicatePreservationPass = (nOut = 3 And hitsCount = 2)
End Function

Function EvalAB(ByVal usePostIngress, ByVal variantId)
  Call ResetState("HARNESS_STRICT")

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap
  Call BuildAmbiguousFixtures(variantId, catAlias, catMap)

  Dim candidateInput
  If CBool(usePostIngress) Then
    Set candidateInput = catAlias
  Else
    candidateInput = catAlias.Keys
  End If

  Dim bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv, ok
  bestKey = "": bestScore = 0
  altKey = "": altScore = 0
  isAmb = False: hasTopTie = False: tieCsv = ""

  ok = HeuristicPickWithAlt("abcd", candidateInput, learn, bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv)

  Dim ret, statOut, conceptOut, acceptedBy
  ret = CBool(ok)
  statOut = ""
  conceptOut = ""
  acceptedBy = "fuzzy/alias"

  If CBool(isAmb) Then
    Call Ambiguity_AddEx("category", "alias", "abcd", "top_tie=[" & tieCsv & "]", "HeuristicPickWithAlt top-score tie fail-closed")
    ret = False
    acceptedBy = "ambiguous/alias"
    statOut = ""
    conceptOut = ""
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
  Set EvalAB = out
End Function

Function EvalSingle(ByVal usePostIngress)
  Call ResetState("OFF")

  Dim learn : Set learn = BuildLearn()
  Dim catAlias, catMap
  Call BuildSingleWinnerFixtures(catAlias, catMap)

  Dim candidateInput
  If CBool(usePostIngress) Then
    Set candidateInput = catAlias
  Else
    candidateInput = catAlias.Keys
  End If

  Dim bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv, ok
  bestKey = "": bestScore = 0
  altKey = "": altScore = 0
  isAmb = False: hasTopTie = False: tieCsv = ""

  ok = HeuristicPickWithAlt("abcd", candidateInput, learn, bestKey, bestScore, altKey, altScore, isAmb, hasTopTie, tieCsv)

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
  Set EvalSingle = out
End Function

Sub EmitFromResult(ByVal r)
  Call EmitResult(CBool(r("ret")), CStr(r("stat")), CStr(r("concept")), CStr(r("is_pitcher")), CStr(r("used_heur")), CStr(r("accepted")), CDbl(r("score")))
End Sub

Sub RunStrictInitGuard()
  G_HARNESS_MODE = "HARNESS_STRICT"
  Set CompilerContext = Nothing
  Set PlanValidationErrors = Nothing

  Call EnsureAmbiguityContextEx(False, "tmp_target10_strict_guard")

  Dim hasKey : hasKey = HasAmbiguousContextInvalid()

  Dim rAbPre, rAbPostA, rAbPostB, rSinglePre, rSinglePost
  Set rAbPre = EvalAB(False, "A")
  Set rAbPostA = EvalAB(True, "A")
  Set rAbPostB = EvalAB(True, "B")
  Set rSinglePre = EvalSingle(False)
  Set rSinglePost = EvalSingle(True)

  Dim parityAB, paritySingle, dupPass, allPass
  parityAB = (CStr(rAbPre("ret")) = CStr(rAbPostA("ret"))) And (CStr(rAbPre("ret")) = CStr(rAbPostB("ret")))
  paritySingle = (CStr(rSinglePre("ret")) = CStr(rSinglePost("ret"))) And (CStr(rSinglePre("stat")) = CStr(rSinglePost("stat")))
  dupPass = DuplicatePreservationPass()

  allPass = ((Not hasKey) And parityAB And paritySingle And dupPass)

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
