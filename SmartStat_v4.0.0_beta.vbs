' SmartStat-Custom-Syntax-Generator_v4.0.0_beta.vbs
Option Explicit
' ================================
' v4.0 Compiler Context (Phase 1)
' ================================
Dim CompilerContext
Dim ApplyPlan
Dim PlanValidationErrors
Dim TRANSACTION_MODE
TRANSACTION_MODE = True

Const DIAG_MODE                = True
Const DIAG_LOG_DIR             = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\"
Const DIAG_PREFIX              = "SmartStat_OperatorDiag_"
Const DIAG_ENV_PREFIX          = "SmartStat_EnvCheck_"
Const SMARTSTAT_DEBUG_FILE     = "SmartStat_Debug.txt"
Const DIAG_MAX_FILE_SIZE_BYTES = "5242880"
Const SMARTSTAT_VERSION = "4.0.0_beta"

' ================================
' v4.0 Phase 3: Harness + Integrity
' ================================
Const HARNESS_ENABLE      = True
Const HARNESS_DIR         = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\Harness\"
Const HARNESS_CONTROL_TAB = "A"      ' Control tabfield (matches existing SMARTSTAT=PLAYER behavior)
Const HARNESS_MODE_CAPTURE_ONLY = "HARNESS_CAPTURE_ONLY"

Dim G_HARNESS_MODE        ' OFF | HARNESS | HARNESS_COMMIT | HARNESS_CAPTURE | HARNESS_CAPTURE_ONLY | HARNESS_STRICT
Dim G_HARNESS_PRE_CP      ' Dict: tabfield -> prior custom prop value
Dim G_HARNESS_PRE_V       ' Dict: tabfield -> prior visible value
Dim G_HARNESS_DIFF_COUNT  ' Integer: number of changed fields detected by Harness_WriteIntegrityDiff
Dim G_HARNESS_POST_CP     ' Dict: tabfield -> post custom prop value
Dim G_HARNESS_POST_V      ' Dict: tabfield -> post visible value
Dim G_HARNESS_POST_SKIPPED_REASON
Dim G_HARNESS_EARLY_EXIT_FINALIZED
Dim G_TRIO_WRITE_ATTEMPTS
Dim G_TRIO_WRITE_SUCCESS_COUNT
Dim G_TRIO_WRITE_FAIL_COUNT
Dim G_TRIO_WRITE_FAILED
Dim G_TRIO_WRITE_LAST_FAIL_TF
Dim G_AMBIGUITY_SUMMARY_EMITTED
Dim G_PHASE_LAST
Dim G_PHASE_FAILED
Dim G_PHASE_ORDER_INDEX

Const PHASE_00_BOOT            = "00.BOOT"
Const PHASE_01_ENV_VALIDATE    = "01.ENV_VALIDATE"
Const PHASE_02_LOAD_CONFIG     = "02.LOAD_CONFIG"
Const PHASE_03_CLASSIFY_FIELDS = "03.CLASSIFY_FIELDS"
Const PHASE_04_APPLY_OVERRIDES = "04.APPLY_OVERRIDES"
Const PHASE_05_DETECT_FILTERS  = "05.DETECT_FILTERS_CATS"
Const PHASE_06_BUILD_OUTMAP    = "06.BUILD_OUTPUT_MAP"
Const PHASE_07_BUILD_SYNTAX    = "07.BUILD_SYNTAX"
Const PHASE_08_PUSH_TO_TRIO    = "08.PUSH_TO_TRIO"
Const PHASE_99_DONE            = "99.DONE"

Dim gDiagRunId, gDiagFile, gFSO, gDiagStarted, gTemplateName, gStartTicks

Function Phase_OrderList()
  Phase_OrderList = Array( _
    PHASE_00_BOOT, _
    PHASE_01_ENV_VALIDATE, _
    PHASE_02_LOAD_CONFIG, _
    PHASE_03_CLASSIFY_FIELDS, _
    PHASE_04_APPLY_OVERRIDES, _
    PHASE_05_DETECT_FILTERS, _
    PHASE_06_BUILD_OUTMAP, _
    PHASE_07_BUILD_SYNTAX, _
    PHASE_08_PUSH_TO_TRIO, _
    PHASE_99_DONE)
End Function

Function Phase_OrderIndex(ByVal phaseId)
  Dim arr, i, p
  p = UCase(Trim(CStr(phaseId)))
  arr = Phase_OrderList()
  For i = LBound(arr) To UBound(arr)
    If UCase(CStr(arr(i))) = p Then
      Phase_OrderIndex = i + 1
      Exit Function
    End If
  Next
  Phase_OrderIndex = 0
End Function

Function Phase_OrderName(ByVal idx)
  Dim arr, z, lo, hi, count, idx0
  arr = Phase_OrderList()
  z = CInt(idx)
  lo = LBound(arr)
  hi = UBound(arr)
  count = hi - lo + 1
  If z < 1 Or z > count Then
    Phase_OrderName = ""
    Exit Function
  End If
  idx0 = lo + (z - 1)
  Phase_OrderName = CStr(arr(idx0))
End Function

Function Ambiguity_HasHits()
  On Error Resume Next
  Ambiguity_HasHits = False
  Call EnsureAmbiguityContext()
  If (CompilerContext Is Nothing) Then Exit Function
  If Not CompilerContext.Exists("ambiguous") Then Exit Function
  If (Not IsObject(CompilerContext("ambiguous"))) Then Exit Function
  If UCase(TypeName(CompilerContext("ambiguous"))) <> "DICTIONARY" Then Exit Function
  Ambiguity_HasHits = (CompilerContext("ambiguous").Count > 0)
  On Error GoTo 0
End Function

Sub Phase_Begin(ByVal phaseId, ByVal msg)
  On Error Resume Next
  Dim gotIdx, expectedPhase
  gotIdx = Phase_OrderIndex(phaseId)

  If gotIdx > 0 Then
    If CLng(G_PHASE_ORDER_INDEX) = 0 Then
      If CBool(DIAG_MODE) Then
        If gotIdx <> 1 Then
          expectedPhase = Phase_OrderName(1)
          Call Diag_WriteLine("PHASE_ORDER_WARN expected=" & expectedPhase & " got=" & CStr(phaseId) & " last=" & CStr(G_PHASE_LAST))
        End If
      End If
      G_PHASE_ORDER_INDEX = gotIdx
    Else
      If gotIdx < CLng(G_PHASE_ORDER_INDEX) Or gotIdx > (CLng(G_PHASE_ORDER_INDEX) + 1) Then
        If CBool(DIAG_MODE) Then
          expectedPhase = Phase_OrderName(CLng(G_PHASE_ORDER_INDEX) + 1)
          If Len(expectedPhase) = 0 Then expectedPhase = "END"
          Call Diag_WriteLine("PHASE_ORDER_WARN expected=" & expectedPhase & " got=" & CStr(phaseId) & " last=" & CStr(G_PHASE_LAST))
        End If
      End If
      G_PHASE_ORDER_INDEX = gotIdx
    End If
  End If

  G_PHASE_LAST = CStr(phaseId)
  Call Diag_WriteHeader(CStr(phaseId), CStr(msg))
  On Error GoTo 0
End Sub

Sub Phase_EndOk(ByVal phaseId, ByVal msg)
  On Error Resume Next
  If Len(Trim(CStr(phaseId))) = 0 Then Exit Sub
  Call Diag_WriteLine("PHASE_END phase=" & CStr(phaseId) & " detail=" & CStr(msg))
  On Error GoTo 0
End Sub

Sub Phase_Fail(ByVal phaseId, ByVal code, ByVal detail, ByVal actionHint)
  On Error Resume Next
  G_PHASE_FAILED = True
  G_PHASE_LAST = CStr(phaseId)
  G_PHASE_ORDER_INDEX = Phase_OrderIndex(phaseId)
  Call Diag_HardFail(CStr(phaseId), CStr(code), CStr(detail), CStr(actionHint))
  If CBool(DIAG_MODE) Then
    If Ambiguity_HasHits() Then Call Diag_WriteAmbiguitySummary()
  End If
  On Error GoTo 0
End Sub

Sub Phase_EarlyExit(ByVal phaseId, ByVal reasonCode, ByVal detail, ByVal actionHint)
  On Error Resume Next
  G_PHASE_FAILED = True
  G_PHASE_LAST = CStr(phaseId)
  G_PHASE_ORDER_INDEX = Phase_OrderIndex(phaseId)
  Call Diag_WriteLine("EARLY_EXIT phase=" & CStr(phaseId) & " code=" & CStr(reasonCode) & " detail=" & CStr(detail) & " action=" & CStr(actionHint))
  If CBool(DIAG_MODE) Then
    If Ambiguity_HasHits() Then Call Diag_WriteAmbiguitySummary()
  End If
  On Error GoTo 0
End Sub

Sub Diag_Init(ByVal templateName)
    If Not DIAG_MODE Then Exit Sub
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    gDiagRunId = Diag_TimestampCompact()
    gTemplateName = templateName
    If Not gFSO.FolderExists(DIAG_LOG_DIR) Then
        On Error Resume Next
        gFSO.CreateFolder DIAG_LOG_DIR
        On Error GoTo 0
    End If
    gDiagFile = DIAG_LOG_DIR & DIAG_PREFIX & gDiagRunId & ".txt"
    gDiagStarted = True
    gStartTicks = Timer

    Diag_WriteHeader PHASE_00_BOOT, "SmartStat Diagnostics Start"
    Diag_WriteLine "RUN_ID=" & gDiagRunId
    Diag_WriteLine "TEMPLATE=" & gTemplateName
    Diag_WriteLine "MACHINE=" & Diag_SafeEnv("COMPUTERNAME") & " USER=" & Diag_SafeEnv("USERNAME")
    Diag_WriteLine "TRIO_ENV=VBScript"
    Diag_WriteLine "----------------------------"
End Sub

Sub Diag_Done()
    If Not gDiagStarted Then Exit Sub
    Diag_Step PHASE_99_DONE, "Completed in " & CStr(Round((Timer - gStartTicks), 3)) & "s"
End Sub

Sub Diag_Step(ByVal phase, ByVal detail)
    If Not DIAG_MODE Then Exit Sub
    Diag_WriteHeader phase, detail
End Sub

Function Diag_Assert(ByVal condition, ByVal phase, ByVal code, ByVal msg, ByVal remedy)
    If condition Then
        Diag_WriteLine "[OK] " & code & ": " & msg
        Diag_Assert = True
        Exit Function
    End If
    Dim fullMsg
    fullMsg = "[FAIL] " & code & ": " & msg
    If Len(remedy) > 0 Then fullMsg = fullMsg & " | Remedy: " & remedy
    Diag_WriteHeader phase, fullMsg
    Diag_OperatorAlert "SmartStat failed at " & phase & " (" & code & "). See log: " & gDiagFile
    Diag_AppendToDebug fullMsg
    Diag_Assert = False
End Function

Sub Diag_OperatorAlert(ByVal message)
    On Error Resume Next
    TrioCmd("gui:error_message " & message)
    On Error GoTo 0
End Sub

Function Diag_FileExistsReadable(ByVal path)
    Diag_FileExistsReadable = False
    If path = "" Then Exit Function
    If Not gFSO.FileExists(path) Then Exit Function
    On Error Resume Next
    Dim f, ok
    ok = False
    Set f = gFSO.OpenTextFile(path, 1, False)
    If Not f Is Nothing Then
        ok = True
        f.Close
    End If
    On Error GoTo 0
    Diag_FileExistsReadable = ok
End Function

Function Diag_SafeEnv(ByVal key)
    On Error Resume Next
    Diag_SafeEnv = ""
    Dim wsh : Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        Diag_SafeEnv = wsh.ExpandEnvironmentStrings("%" & key & "%")
    End If
    On Error GoTo 0
    If IsNull(Diag_SafeEnv) Then Diag_SafeEnv = ""
End Function

Function Diag_RegRead(ByVal regPath, ByRef valueOut)
    On Error Resume Next
    Dim wsh, v : Set wsh = CreateObject("WScript.Shell")
    Err.Clear
    v = wsh.RegRead(regPath)
    If Err.Number <> 0 Then
        Diag_RegRead = False
        valueOut = ""
        Err.Clear
    Else
        Diag_RegRead = True
        valueOut = v
    End If
    On Error GoTo 0
End Function

Function Diag_CheckWSHEnabled()
    Dim val, ok
    ok = Diag_RegRead("HKLM\Software\Microsoft\Windows Script Host\Settings\Enabled", val)
    If ok Then
        If IsNumeric(val) Then
            Diag_CheckWSHEnabled = (CLng(val) <> 0)
            Exit Function
        End If
    End If
    Diag_CheckWSHEnabled = True
End Function

Function Diag_CheckADOAvailable()
    Dim clsid, ok
    ok = Diag_RegRead("HKCR\ADODB.Stream\CLSID", clsid)
    Diag_CheckADOAvailable = ok And (Len(clsid) > 0)
End Function

Function Diag_CheckWriteAccess(ByVal folderPath)
    Dim probe, ok
    ok = False
    probe = folderPath & DIAG_ENV_PREFIX & gDiagRunId & ".tmp"
    On Error Resume Next
    Dim ts : Set ts = gFSO.OpenTextFile(probe, 2, True)
    If Not ts Is Nothing Then
        ts.WriteLine "probe"
        ts.Close
        gFSO.DeleteFile probe, True
        ok = True
    End If
    On Error GoTo 0
    Diag_CheckWriteAccess = ok
End Function

Sub Diag_WriteHeader(ByVal phase, ByVal title)
    If Not DIAG_MODE Then Exit Sub
    Diag_WriteLine ">>> [" & phase & "] " & title
End Sub

Sub Diag_WriteLine(ByVal s)
    If Not DIAG_MODE Then Exit Sub
    Dim line : line = "[" & Diag_TimestampHuman() & "] " & s
    On Error Resume Next
    If gFSO.FileExists(gDiagFile) Then
        If gFSO.GetFile(gDiagFile).Size > DIAG_MAX_FILE_SIZE_BYTES Then
            gFSO.DeleteFile gDiagFile, True
        End If
    End If
    Dim ts : Set ts = gFSO.OpenTextFile(gDiagFile, 8, True)
    If Not ts Is Nothing Then
        ts.WriteLine line
        ts.Close
    End If
    On Error GoTo 0
End Sub

' ==========================================
' v4.0 Phase 2: Ambiguity recording
' ==========================================
Sub Ambiguity_Add(where, kind, inputTxt, bestKey, bestScore, altKey, altScore)
  On Error Resume Next

  If (CompilerContext Is Nothing) Then Exit Sub
  If Not CompilerContext.Exists("ambiguous") Then Exit Sub

  Dim d: Set d = CompilerContext("ambiguous")
  Dim id: id = kind & "|" & where & "|" & UCase(Trim(CStr(inputTxt)))

  Dim msg
  msg = kind & " ambiguous in " & where & " input=[" & CStr(inputTxt) & _
        "] best=[" & bestKey & "](" & ScoreStr(bestScore) & _
        ") alt=[" & altKey & "](" & ScoreStr(altScore) & ")"

  d(id) = msg
  Call Diag_WriteLine("AMBIGUITY: " & msg)

  On Error GoTo 0
End Sub

Sub Ambiguity_AddEx(field, phase, inputValue, candidates, note)
  On Error Resume Next

  If (CompilerContext Is Nothing) Then Exit Sub
  If Not CompilerContext.Exists("ambiguous") Then
    Dim initAmb: Set initAmb = CreateObject("Scripting.Dictionary")
    Set CompilerContext("ambiguous") = initAmb
  ElseIf (Not IsObject(CompilerContext("ambiguous"))) Or (UCase(TypeName(CompilerContext("ambiguous"))) <> "DICTIONARY") Then
    Call Diag_WriteLine("AMBIGUITY: ambiguous context reset in Ambiguity_AddEx (TypeName=" & TypeName(CompilerContext("ambiguous")) & ")")
    Dim resetAmb: Set resetAmb = CreateObject("Scripting.Dictionary")
    Set CompilerContext("ambiguous") = resetAmb
  End If

  Dim d: Set d = CompilerContext("ambiguous")
  If (d Is Nothing) Then Exit Sub

  Dim f: f = LCase(Trim(CStr(field)))
  Dim p: p = LCase(Trim(CStr(phase)))
  Dim inp: inp = UCase(Trim(CStr(inputValue)))
  Dim cand: cand = Trim(CStr(candidates))
  Dim n: n = Trim(CStr(note))

  Dim id
  id = f & "|" & p & "|" & inp & "|" & UCase(cand)

  If Not d.Exists(id) Then
    Dim msg
    msg = phase & " ambiguous in " & field & " input=[" & CStr(inputValue) & _
          "] candidates=[" & CStr(candidates) & "] note=[" & CStr(n) & "]"

    d(id) = msg
    Call Diag_WriteLine("AMBIGUITY: " & msg)
  End If

  On Error GoTo 0
End Sub

Sub Diag_AppendToDebug(ByVal s)
    On Error Resume Next
    Dim path : path = DIAG_LOG_DIR & SMARTSTAT_DEBUG_FILE
	Diag_TrimLogFileSize path
    Dim ts : Set ts = gFSO.OpenTextFile(path, 8, True)
    If Not ts Is Nothing Then
        ts.WriteLine "[" & Diag_TimestampHuman() & "] " & s
        ts.Close
    End If
    On Error GoTo 0
End Sub

Function Diag_TimestampCompact()
    Dim d : d = Now
    Diag_TimestampCompact = Right("0000" & Year(d), 4) & Right("00" & Month(d), 2) & Right("00" & Day(d), 2) & "_" & Right("00" & Hour(d), 2) & Right("00" & Minute(d), 2) & Right("00" & Second(d), 2)
End Function

Function Diag_TimestampHuman()
    Dim d : d = Now
    Diag_TimestampHuman = Right("0000" & Year(d), 4) & "-" & Right("00" & Month(d), 2) & "-" & Right("00" & Day(d), 2) & " " & Right("00" & Hour(d), 2) & ":" & Right("00" & Minute(d), 2) & ":" & Right("00" & Second(d), 2)
End Function

Function Diag_Check_Environment()
    If Not DIAG_MODE Then Diag_Check_Environment = True : Exit Function
    Call Phase_Begin(PHASE_01_ENV_VALIDATE, "Validating environment and prerequisites")
    Dim ok
    ok = Diag_Assert(Diag_CheckWSHEnabled(), PHASE_01_ENV_VALIDATE, "ENV.WSH", "Windows Script Host appears enabled", "Enable: HKLM\Software\Microsoft\Windows Script Host\Settings\Enabled=1 (or key absent)")
    If Not ok Then
        Call Phase_EarlyExit(PHASE_01_ENV_VALIDATE, "ENV.WSH_FAIL", "Windows Script Host check failed.", "Enable WSH and rerun SmartStat.")
        Diag_Check_Environment = False
        Exit Function
    End If
    ok = Diag_Assert(Diag_CheckWriteAccess(DIAG_LOG_DIR), PHASE_01_ENV_VALIDATE, "ENV.LOGWRITE", "Write access to " & DIAG_LOG_DIR, "Grant write perms to " & DIAG_LOG_DIR & " or choose a writable folder")
    If Not ok Then
        Call Phase_EarlyExit(PHASE_01_ENV_VALIDATE, "ENV.LOGWRITE_FAIL", "Cannot write to diagnostics directory.", "Grant write permissions to SmartStat\\DiagLogs and retry.")
        Diag_Check_Environment = False
        Exit Function
    End If
    Dim adoOK : adoOK = Diag_CheckADOAvailable()
    If adoOK Then
        Diag_WriteLine "[OK] ENV.ADO: ADODB present"
    Else
        Diag_WriteLine "[WARN] ENV.ADO: ADODB registry not found; if SmartStat uses ADO on this page, install MDAC/ADO."
    End If
    Call Phase_EndOk(PHASE_01_ENV_VALIDATE, "Environment validation passed")
    Diag_Check_Environment = True
End Function

Function Diag_Check_ConfigPresence(ByVal mappingsPath, ByVal overridesPath, ByVal templateCfgPath)
    If Not DIAG_MODE Then Diag_Check_ConfigPresence = True : Exit Function
    Call Phase_Begin(PHASE_02_LOAD_CONFIG, "Probing config files")
    Dim ok
    ok = Diag_Assert(Diag_FileExistsReadable(mappingsPath), PHASE_02_LOAD_CONFIG, "CFG.MAP", "Mappings file readable: " & mappingsPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then
        Call Phase_EarlyExit(PHASE_02_LOAD_CONFIG, "CFG.MAP_FAIL", "Mappings file is missing/unreadable: " & mappingsPath, "Restore SmartStat_Mappings.ini and verify read permissions.")
        Diag_Check_ConfigPresence = False
        Exit Function
    End If
    ok = Diag_Assert(Diag_FileExistsReadable(overridesPath), PHASE_02_LOAD_CONFIG, "CFG.OVR", "Static overrides file readable: " & overridesPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then
        Call Phase_EarlyExit(PHASE_02_LOAD_CONFIG, "CFG.OVR_FAIL", "Static overrides file is missing/unreadable: " & overridesPath, "Restore SmartStat_StaticOverrides.ini and verify read permissions.")
        Diag_Check_ConfigPresence = False
        Exit Function
    End If
    ok = Diag_Assert(Diag_FileExistsReadable(templateCfgPath), PHASE_02_LOAD_CONFIG, "CFG.TPL", "Template config file readable: " & templateCfgPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then
        Call Phase_EarlyExit(PHASE_02_LOAD_CONFIG, "CFG.TPL_FAIL", "Template config file is missing/unreadable: " & templateCfgPath, "Restore SmartStat_TemplateConfig.ini and verify read permissions.")
        Diag_Check_ConfigPresence = False
        Exit Function
    End If
    Call Phase_EndOk(PHASE_02_LOAD_CONFIG, "Configuration files verified")
    Diag_Check_ConfigPresence = True
End Function

Sub Diag_Mark_LoadConfig(ByVal detail)     : Diag_Step PHASE_02_LOAD_CONFIG, detail     : End Sub
Sub Diag_Mark_Classify(ByVal detail)       : Diag_Step PHASE_03_CLASSIFY_FIELDS, detail : End Sub
Sub Diag_Mark_ApplyOverrides(ByVal detail) : Diag_Step PHASE_04_APPLY_OVERRIDES, detail : End Sub
Sub Diag_Mark_DetectFilters(ByVal detail)  : Diag_Step PHASE_05_DETECT_FILTERS, detail  : End Sub
Sub Diag_Mark_BuildOutMap(ByVal detail)    : Diag_Step PHASE_06_BUILD_OUTMAP, detail    : End Sub
Sub Diag_Mark_BuildSyntax(ByVal detail)    : Diag_Step PHASE_07_BUILD_SYNTAX, detail    : End Sub
Sub Diag_Mark_PushToTrio(ByVal detail)     : Diag_Step PHASE_08_PUSH_TO_TRIO, detail    : End Sub

Sub Diag_HardFail(ByVal phase, ByVal code, ByVal msg, ByVal remedy)
    Call Diag_Assert(False, phase, code, msg, remedy)
End Sub


' --- Global league tag for name adjustments ---
Dim G_SPORT_TAG: G_SPORT_TAG = "MLB"

' Precompiled regex for trimming redundant season chaining segments
Dim G_REGEX_SEASON_TRAILING_SEGMENT
Set G_REGEX_SEASON_TRAILING_SEGMENT = Nothing

' =========================
' - Robust category resolver (direct, alias, doubled-letter rescue, one-edit + transposition for short tokens)
' - Qualifier + filter_tabfields resolution with fuzzy fallback
' - Per-row filter application for output_map (no stacking all filters into one path)
' - Learn logging for unresolved categories and qualifiers/filters
' - StaticOverrides respected
' - VBScript-safe indexing (parentheses), CreateObject for RegExp
' =========================

Dim startTime: startTime = Timer
Call Main()

Sub Main()
  On Error Resume Next
  G_PHASE_LAST = ""
  G_PHASE_FAILED = False
  G_PHASE_ORDER_INDEX = 0
  Dim x_tmplForDiag: x_tmplForDiag = TrioCmd("page:getpagetemplate")
  Call Diag_Init(x_tmplForDiag)

  Call Diag_Log("VERSION=" & SMARTSTAT_VERSION)
  Dim LOG_FILE:      LOG_FILE      = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\SmartStat_LearnDebug.txt"

  ' ================================
  ' v4.0 Initialize Compiler Context
  ' ================================
  Set CompilerContext = CreateObject("Scripting.Dictionary")
  Set ApplyPlan = CreateObject("Scripting.Dictionary")
  Set PlanValidationErrors = CreateObject("Scripting.Dictionary")

  G_TRIO_WRITE_ATTEMPTS = 0
  G_TRIO_WRITE_SUCCESS_COUNT = 0
  G_TRIO_WRITE_FAIL_COUNT = 0
  G_TRIO_WRITE_FAILED = False
  G_TRIO_WRITE_LAST_FAIL_TF = ""
  G_AMBIGUITY_SUMMARY_EMITTED = False
  G_HARNESS_EARLY_EXIT_FINALIZED = False

  ' v4.0 Phase 2: ambiguity & confidence context
  Call EnsureAmbiguityContext()
  If Not CompilerContext.Exists("ambiguous") Then
    Dim AmbiguityHits: Set AmbiguityHits = CreateObject("Scripting.Dictionary")
    Set CompilerContext("ambiguous") = AmbiguityHits
  ElseIf (Not IsObject(CompilerContext("ambiguous"))) Or (UCase(TypeName(CompilerContext("ambiguous"))) <> "DICTIONARY") Then
    Call Diag_WriteLine("TX: ambiguous context reset in Main (TypeName=" & TypeName(CompilerContext("ambiguous")) & ")")
    Dim AmbiguityHitsReset: Set AmbiguityHitsReset = CreateObject("Scripting.Dictionary")
    Set CompilerContext("ambiguous") = AmbiguityHitsReset
  End If

  ' ================================
  ' v4.0 Phase 3: Harness bootstrap
  ' ================================
  G_HARNESS_MODE = "OFF"
  If HARNESS_ENABLE Then
    G_HARNESS_MODE = Harness_GetModeFromControlA()
    If G_HARNESS_MODE <> "OFF" Then
      Call Harness_EnsureHarnessDir()
      Set G_HARNESS_PRE_CP = CreateTextDict()
      Set G_HARNESS_PRE_V  = CreateTextDict()
      Call Harness_SnapshotPageState(G_HARNESS_PRE_V, G_HARNESS_PRE_CP)
      Call Harness_ResetPostState()
      If G_HARNESS_MODE = HARNESS_MODE_CAPTURE_ONLY Then
        Call Harness_MarkPostSkipped("CAPTURE_ONLY")
        Call Harness_WriteSnapshotArtifact("CAPTURE_ONLY")
        Call Harness_WriteGroupedDiffArtifact("CAPTURE_ONLY")
        Call Diag_WriteLine("HARNESS: capture-only mode; exiting before pipeline")
        Call Diag_WriteLine("TX: EARLY EXIT - HARNESS_CAPTURE_ONLY")
        Call Diag_Done()
        FinalizeAndRefresh LOG_FILE, startTime
        Exit Sub
      End If
      Call Diag_WriteLine("HARNESS: mode=" & G_HARNESS_MODE & " preSnapshotTabs=" & CStr(G_HARNESS_PRE_CP.Count))

      If G_HARNESS_MODE = "HARNESS_CAPTURE" Then
        Call Harness_WriteFixtureFile(G_HARNESS_PRE_V, G_HARNESS_PRE_CP, x_tmplForDiag)
        Call Diag_WriteLine("HARNESS: capture complete (no pipeline executed)")
        Call Harness_CapturePostSnapshot()
        Call Harness_WriteSnapshotArtifact("HARNESS_CAPTURE")
        Call Harness_WriteGroupedDiffArtifact("HARNESS_CAPTURE")
        Call Diag_WriteLine("TX: EARLY EXIT - HARNESS_CAPTURE")
        Call Diag_Done()
        FinalizeAndRefresh LOG_FILE, startTime
        Exit Sub
      End If
    End If
  End If

  LOG_FILE      = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\SmartStat_LearnDebug.txt"
  If Not Diag_Check_Environment() Then
    Call Harness_FinalizeEarlyExitArtifacts("ENV_VALIDATE_FAIL")
    Call Diag_WriteLine("TX: EARLY EXIT - ENV_VALIDATE_FAIL")
    Call Diag_Done()
    Call FinalizeAndRefresh(LOG_FILE, startTime)
    Exit Sub
  End If
  Dim SRC_DIR:       SRC_DIR       = "E:\EDRIVE\UNIVERSAL\SmartStat\"
  ' --- Auto-handle Google Drive nested folder case (SmartStat\SmartStat) ---
  If Not gFSO.FileExists(SRC_DIR & "SmartStat_TemplateConfig.ini") Then
    If gFSO.FolderExists(SRC_DIR & "SmartStat\") And _
      gFSO.FileExists(SRC_DIR & "SmartStat\SmartStat_TemplateConfig.ini") Then
      SRC_DIR = SRC_DIR & "SmartStat\"
      Diag_WriteLine "[WARN] Adjusted SRC_DIR to nested SmartStat folder: " & SRC_DIR
    End If
  End If

  Dim MAPPINGS_INI:  MAPPINGS_INI  = SRC_DIR & "SmartStat_Mappings.ini"
  Dim LEARN_INI:     LEARN_INI     = SRC_DIR & "SmartStat_Mappings.learn.ini"

  ' --- Sport-aware mappings selection (MLB default; supports NHL/NBA/etc) ---
  Dim sport_tag: sport_tag = UCase(Trim(TrioCmd("trio:get_global_variable league")))
  If Len(sport_tag) = 0 Then sport_tag = UCase(Trim(GetEnv("SMARTSTAT_SPORT")))
  If Len(sport_tag) > 0 And sport_tag <> "MLB" Then
    Dim sMap, sLearn
    sMap   = ResolveSportMappingsPath(SRC_DIR, sport_tag, False)   ' e.g., SmartStat_MappingsNHL.ini or MappingsNHL.ini
    sLearn = ResolveSportMappingsPath(SRC_DIR, sport_tag, True)    ' e.g., SmartStat_MappingsNHL.learn.ini or MappingsNHL.learn.ini
    If Len(sMap) > 0 Then MAPPINGS_INI = sMap
    If Len(sLearn) > 0 Then LEARN_INI   = sLearn
  End If
  If Len(sport_tag) = 0 Then sport_tag = "MLB"
  G_SPORT_TAG = sport_tag

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(MAPPINGS_INI) Then
    Dim altMappings: altMappings = ResolveMappingsPath()
    If Len(altMappings) > 0 Then
      MAPPINGS_INI = altMappings
      LEARN_INI = fso.BuildPath(fso.GetParentFolderName(MAPPINGS_INI), "SmartStat_Mappings.learn.ini")
      SRC_DIR = fso.GetParentFolderName(MAPPINGS_INI) & "\"
    End If
  End If

  If Not Diag_Check_ConfigPresence(MAPPINGS_INI, SRC_DIR & "SmartStat_StaticOverrides.ini", SRC_DIR & "SmartStat_TemplateConfig.ini") Then
    Call Harness_FinalizeEarlyExitArtifacts("CONFIG_PRESENCE_FAIL")
    Call Diag_WriteLine("TX: EARLY EXIT - CONFIG_PRESENCE_FAIL")
    Call Diag_Done()
    Call FinalizeAndRefresh(LOG_FILE, startTime)
    Exit Sub
  End If

  Diag_Mark_LoadConfig "Loading INI mappings"
  EnsureLogDir LOG_FILE
  EnsureLogDir DIAG_LOG_DIR & "._probe"
  LogLine LOG_FILE, "=== v4.0.0_beta START ==="
  LogLine LOG_FILE, "INFO: Using mappings at [" & MAPPINGS_INI & "]"
  LogLine LOG_FILE, "INFO: Learn file path  [" & LEARN_INI & "]"

  Dim ini: Set ini = LoadIni(MAPPINGS_INI)
  If ini Is Nothing Then
    GuiErr "Failed to load INI: " & MAPPINGS_INI
    LogLine LOG_FILE, "ERROR: INI load failed (" & MAPPINGS_INI & ")"

    ' still run so static overrides etc. can apply
    ExecuteTemplatePipeline SRC_DIR, Nothing, LEARN_INI, NewTextDict(), NewTextDict(), NewTextDict()

    ' v4.0 Phase 3: integrity diff (even in ini-fail path, if harness active)
    If HARNESS_ENABLE Then
      If G_HARNESS_MODE <> "OFF" Then
        If Not (G_HARNESS_PRE_CP Is Nothing) Then
          Call Harness_WriteIntegrityDiff(G_HARNESS_PRE_CP, ApplyPlan)
        End If
      End If
    End If

    If TRANSACTION_MODE Then
      Call Diag_Log("TRANSACTION_MODE=" & CStr(TRANSACTION_MODE))
      Err.Clear

      ' Harness default safety: do NOT commit unless explicitly allowed
      If G_HARNESS_MODE = "HARNESS" Then
        Call Diag_WriteLine("HARNESS: no-commit mode; skipping Stage_CommitTransaction (ini-load path)")
      ElseIf G_HARNESS_MODE = "HARNESS_STRICT" Then
        If CInt(G_HARNESS_DIFF_COUNT) > 0 Then
          Call Diag_WriteLine("HARNESS_STRICT: non-empty diff (" & CStr(G_HARNESS_DIFF_COUNT) & "); blocking commit (ini-load path)")
          Call Diag_OperatorAlert("SmartStat Harness Strict: Diff detected. Commit blocked.")
        Else
          Call Diag_WriteLine("HARNESS_STRICT: precommit validate/commit path entered (diffCount=" & CStr(G_HARNESS_DIFF_COUNT) & ")")
          Dim txOkIniS: txOkIniS = Stage_ValidatePlan()
          If Err.Number <> 0 Then
            Call Diag_WriteLine("TX: Stage_ValidatePlan runtime error (ini-load path) Err.Number=" & CStr(Err.Number) & " Err.Description=" & CStr(Err.Description))
            Err.Clear
          End If

          If txOkIniS Then
            Call Stage_CommitTransaction()
          Else
            Call Diag_OperatorAlert("SmartStat aborted: Validation failure. No changes applied.")
            Call Diag_Log("VALIDATION FAILURE - Transaction Aborted")
          End If
        End If
      Else
        Dim txOkIni: txOkIni = Stage_ValidatePlan()
        If Err.Number <> 0 Then
          Call Diag_WriteLine("TX: Stage_ValidatePlan runtime error (ini-load path) Err.Number=" & CStr(Err.Number) & " Err.Description=" & CStr(Err.Description))
          Err.Clear
        End If

        If txOkIni Then
          Call Stage_CommitTransaction()
        Else
          Call Diag_OperatorAlert("SmartStat aborted: Validation failure. No changes applied.")
          Call Diag_Log("VALIDATION FAILURE - Transaction Aborted")
        End If
      End If
    End If

    ' finalize and return (no GoTo)
    Call Diag_WriteLine("TX: EARLY EXIT - MAPPINGS_INI_LOAD_FAIL")
    Call Diag_Done()
    Call FinalizeAndRefresh(LOG_FILE, startTime)
    Exit Sub
  End If

  Dim transforms:   Set transforms   = LoadTransforms(ini)
  Dim rxTransforms: Set rxTransforms = LoadRegexDict(ini, "TRANSFORMS_REGEX")
  Dim learn:        Set learn        = LoadLearnKnobs(ini)

  CompilerContext("learn") = learn

  ExecuteTemplatePipeline SRC_DIR, ini, LEARN_INI, transforms, rxTransforms, learn

  Dim txPostApplyCount: txPostApplyCount = -1
  If Not (ApplyPlan Is Nothing) Then txPostApplyCount = ApplyPlan.Count
  Call Diag_WriteLine("TX: POST_PIPELINE TRANSACTION_MODE=" & CStr(TRANSACTION_MODE) & " (TypeName=" & TypeName(TRANSACTION_MODE) & ") DIAG_MODE=" & CStr(DIAG_MODE) & " HARNESS=" & CStr(G_HARNESS_MODE) & " ApplyPlan.Count=" & CStr(txPostApplyCount))

  ' v4.0 Phase 3: integrity diff (success path)
  If HARNESS_ENABLE Then
    If G_HARNESS_MODE <> "OFF" Then
      If Not (G_HARNESS_PRE_CP Is Nothing) Then
        Call Harness_WriteIntegrityDiff(G_HARNESS_PRE_CP, ApplyPlan)
      End If
    End If
  End If

  If TRANSACTION_MODE Then
    Call Diag_WriteLine("TX: ENTER_TRANSACTION_BLOCK")
    Err.Clear

    ' Harness default safety: do NOT commit unless explicitly allowed
    If G_HARNESS_MODE = "HARNESS" Then
      Call Diag_WriteLine("HARNESS: no-commit mode; skipping Stage_CommitTransaction (main path)")
    ElseIf G_HARNESS_MODE = "HARNESS_STRICT" Then
      If CInt(G_HARNESS_DIFF_COUNT) > 0 Then
        Call Diag_WriteLine("HARNESS_STRICT: non-empty diff (" & CStr(G_HARNESS_DIFF_COUNT) & "); blocking commit (main path)")
        Call Diag_OperatorAlert("SmartStat Harness Strict: Diff detected. Commit blocked.")
      Else
        Call Diag_WriteLine("HARNESS_STRICT: precommit validate/commit path entered (diffCount=" & CStr(G_HARNESS_DIFF_COUNT) & ")")
        Dim txPlanCountS: txPlanCountS = -1
        If Not (ApplyPlan Is Nothing) Then txPlanCountS = ApplyPlan.Count
        Call Diag_WriteLine("TX: ENTER_VALIDATE (ApplyPlan.Count=" & CStr(txPlanCountS) & ")")
        Dim txOkS: txOkS = Stage_ValidatePlan()
        Dim txErrNumS, txErrDescS, txPvCountS
        txErrNumS = Err.Number
        txErrDescS = CStr(Err.Description)
        If (PlanValidationErrors Is Nothing) Then
          txPvCountS = -1
        Else
          txPvCountS = PlanValidationErrors.Count
        End If
        Call Diag_WriteLine("TX: VALIDATE_RESULT=" & CStr(txOkS) & " Err=" & CStr(txErrNumS) & ":" & txErrDescS & " PlanValidationErrors.Count=" & CStr(txPvCountS))

        If txErrNumS <> 0 Then
          Call Diag_WriteLine("TX: Stage_ValidatePlan runtime error (main path) Err.Number=" & CStr(txErrNumS) & " Err.Description=" & txErrDescS)
          Err.Clear
        End If

        If txOkS Then
          Call Diag_WriteLine("TX: ENTER_COMMIT (ApplyPlan.Count=" & CStr(txPlanCountS) & ")")
          Call Stage_CommitTransaction()
        Else
          Call Diag_OperatorAlert("SmartStat aborted: Validation failure. No changes applied.")
          If (PlanValidationErrors Is Nothing) Then
            Call Diag_WriteLine("TX: PlanValidationErrors = Nothing")
          Else
            Call Diag_WriteLine("TX: PlanValidationErrors.Count = " & CStr(PlanValidationErrors.Count))
          End If
        End If
        If (PlanValidationErrors Is Nothing) Then
          Call Diag_WriteLine("TX: PlanValidationErrors = Nothing")
        Else
          Call Diag_WriteLine("TX: PlanValidationErrors.Count = " & CStr(PlanValidationErrors.Count))
        End If

        If Not (PlanValidationErrors Is Nothing) Then
          If PlanValidationErrors.Count > 0 Then
            Call Diag_WriteLine("TX: VALIDATION ERRORS:")

            Dim txErrKeys: txErrKeys = PlanValidationErrors.Keys
            Dim txErrI, txErrJ, txErrTmp, txErrKey
            For txErrI = 0 To UBound(txErrKeys) - 1
              For txErrJ = txErrI + 1 To UBound(txErrKeys)
                If StrComp(CStr(txErrKeys(txErrI)), CStr(txErrKeys(txErrJ)), vbTextCompare) > 0 Then
                  txErrTmp = txErrKeys(txErrI)
                  txErrKeys(txErrI) = txErrKeys(txErrJ)
                  txErrKeys(txErrJ) = txErrTmp
                End If
              Next
            Next

            For txErrI = 0 To UBound(txErrKeys)
              txErrKey = CStr(txErrKeys(txErrI))
              Call Diag_WriteLine("TX:   " & txErrKey & " = " & CStr(PlanValidationErrors(txErrKey)))
            Next
          End If
        End If

        If Not txOkS Then
          Call Diag_WriteLine("TX: VALIDATION FAILURE - Transaction Aborted")
        End If
      End If
    Else
      Dim txPlanCount: txPlanCount = -1
      If Not (ApplyPlan Is Nothing) Then txPlanCount = ApplyPlan.Count
      Call Diag_WriteLine("TX: ENTER_VALIDATE (ApplyPlan.Count=" & CStr(txPlanCount) & ")")
      Dim txOk: txOk = Stage_ValidatePlan()
      Dim txErrNum, txErrDesc, txPvCount
      txErrNum = Err.Number
      txErrDesc = CStr(Err.Description)
      If (PlanValidationErrors Is Nothing) Then
        txPvCount = -1
      Else
        txPvCount = PlanValidationErrors.Count
      End If
      Call Diag_WriteLine("TX: VALIDATE_RESULT=" & CStr(txOk) & " Err=" & CStr(txErrNum) & ":" & txErrDesc & " PlanValidationErrors.Count=" & CStr(txPvCount))

      If txErrNum <> 0 Then
        Call Diag_WriteLine("TX: Stage_ValidatePlan runtime error (main path) Err.Number=" & CStr(txErrNum) & " Err.Description=" & txErrDesc)
        Err.Clear
      End If

      If txOk Then
        Call Diag_WriteLine("TX: ENTER_COMMIT (ApplyPlan.Count=" & CStr(txPlanCount) & ")")
        Call Stage_CommitTransaction()
      Else
        Call Diag_OperatorAlert("SmartStat aborted: Validation failure. No changes applied.")
        If (PlanValidationErrors Is Nothing) Then
          Call Diag_WriteLine("TX: PlanValidationErrors = Nothing")
        Else
          Call Diag_WriteLine("TX: PlanValidationErrors.Count = " & CStr(PlanValidationErrors.Count))
        End If
        Call Diag_WriteLine("TX: VALIDATION FAILURE - Transaction Aborted")
      End If
    End If
  Else
    Call Diag_WriteLine("TX: TRANSACTION_MODE=False; skipping transaction commit block")
  End If

  If HARNESS_ENABLE Then
    If UCase(Trim(CStr(G_HARNESS_MODE))) <> "OFF" Then
      Call Harness_CapturePostSnapshot()
      Call Harness_WriteSnapshotArtifact("POST_PIPELINE")
      Call Harness_WriteGroupedDiffArtifact("POST_PIPELINE")
    End If
  End If

  ' normal finalize (no label)
  Call Diag_Done()
  Call FinalizeAndRefresh(LOG_FILE, startTime)
End Sub

Sub FinalizeAndRefresh(logFile, startT)
  Dim elapsed: elapsed = Round(Timer - startT, 2)
  LogLine logFile, "SCRIPT RUNTIME: " & elapsed & " seconds"
  LogLine logFile, "=== v4.0.0_beta END ==="
  LogLine logFile, " "
  On Error GoTo 0
  Call SmartStat_RefreshSocketData()
End Sub

' ==========================================
' v4.0 Stage 4 � Structural Validation Gate
' - Allows clears ("") and control strings (e.g., SMARTSTAT=PLAYER)
' - Only enforces moustache pairing when moustaches are present
' ==========================================
Function Stage_ValidatePlan()
  Stage_ValidatePlan = True

  Call EnsureAmbiguityContext()
  If Not (CompilerContext Is Nothing) Then
    If CompilerContext.Exists("ambiguous") Then
      If IsObject(CompilerContext("ambiguous")) Then
        If UCase(TypeName(CompilerContext("ambiguous"))) = "DICTIONARY" Then
          If CompilerContext("ambiguous").Count > 0 Then
            Call Diag_WriteAmbiguitySummary()
          End If
        End If
      End If
    End If
  End If

  If ApplyPlan Is Nothing Then
    Call Diag_WriteLine("TX: ApplyPlan is Nothing")
    Stage_ValidatePlan = False
    Call Diag_WriteLine("TX: EARLY EXIT - APPLYPLAN_NOTHING")
    If (PlanValidationErrors Is Nothing) Then
      Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
    End If
    If PlanValidationErrors.Count = 0 Then
      PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
    End If
    Exit Function
  End If

    ' v4.0 Phase 2: hard block if qualifier failed resolution
  If Not (PlanValidationErrors Is Nothing) Then
    If PlanValidationErrors.Exists("QUALIFIER_UNRESOLVED") Then
      Call Diag_WriteLine("TX: QUALIFIER_UNRESOLVED - blocking apply")
      Stage_ValidatePlan = False
      Call Diag_WriteLine("TX: EARLY EXIT - QUALIFIER_UNRESOLVED_PRECHECK")
      If (PlanValidationErrors Is Nothing) Then
        Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
      End If
      If PlanValidationErrors.Count = 0 Then
        PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
      End If
      Exit Function
    End If
  End If

  ' v4.0 Phase 2: block commit if ambiguity exists (unless explicitly allowed)
  If Not (CompilerContext Is Nothing) Then
    If CompilerContext.Exists("ambiguous") Then
      Dim ambTypeName: ambTypeName = TypeName(CompilerContext("ambiguous"))
      If (Not IsObject(CompilerContext("ambiguous"))) Or UCase(CStr(ambTypeName)) <> "DICTIONARY" Then
        Call Diag_WriteLine("TX: AMBIGUITY_GATE ambiguous context invalid TypeName=" & CStr(ambTypeName))
        Call Diag_WriteLine("TX: PHASE_CODE=AMBIGUOUS_GATE")
        Call Diag_WriteLine("TX: EARLY EXIT - AMBIGUOUS_CONTEXT_INVALID")
        Call Diag_WriteLine("TX: AMBIGUITY_GATE_SUMMARY count=0 keys=[]")
        If (Not IsObject(PlanValidationErrors)) Then
          Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
        ElseIf (PlanValidationErrors Is Nothing) Then
          Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
        End If
        PlanValidationErrors.RemoveAll
        PlanValidationErrors("AMBIGUOUS_CONTEXT_INVALID") = "CompilerContext('ambiguous') is not a Dictionary. Apply blocked to prevent unsafe commit."
        Stage_ValidatePlan = False
        Call Diag_WriteLine("Ambiguity detected. No changes applied.")
        If Err.Number <> 0 Then Err.Clear
        Exit Function
      End If

      Dim amb: Set amb = CompilerContext("ambiguous")
      If amb.Count > 0 Then
        Dim allowAmb: allowAmb = False
        If CompilerContext.Exists("learn") Then
          Dim learnTypeName: learnTypeName = TypeName(CompilerContext("learn"))
          If IsObject(CompilerContext("learn")) And UCase(CStr(learnTypeName)) = "DICTIONARY" Then
            Dim lk: Set lk = CompilerContext("learn")
            If lk.Exists("allow_ambiguous_apply") Then allowAmb = CBool(lk("allow_ambiguous_apply"))
          Else
            Call Diag_WriteLine("TX: AMBIGUITY_GATE learn context invalid TypeName=" & CStr(learnTypeName))
          End If
        End If

        If Not allowAmb Then
          Stage_ValidatePlan = False
          PlanValidationErrors.RemoveAll
          PlanValidationErrors("AMBIGUOUS") = "Ambiguous mapping detected; operator choice required."
          Call Diag_WriteLine("TX: Ambiguity gate blocked apply (count=" & CStr(amb.Count) & ")")
          Call Diag_WriteLine("TX: EARLY EXIT - AMBIGUOUS_GATE")
          Call Diag_WriteLine("TX: PHASE_CODE=AMBIGUOUS_GATE")
          Dim ambKeys, ambI, ambJ, ambTmp, ambKey, ambKeysCsv, ambShown, ambTotal
          ambKeys = amb.Keys
          If IsArray(ambKeys) Then
            For ambI = 0 To UBound(ambKeys) - 1
              For ambJ = ambI + 1 To UBound(ambKeys)
                If StrComp(CStr(ambKeys(ambI)), CStr(ambKeys(ambJ)), vbTextCompare) > 0 Then
                  ambTmp = ambKeys(ambI)
                  ambKeys(ambI) = ambKeys(ambJ)
                  ambKeys(ambJ) = ambTmp
                End If
              Next
            Next

            ambKeysCsv = ""
            For ambI = 0 To UBound(ambKeys)
              If Len(ambKeysCsv) > 0 Then ambKeysCsv = ambKeysCsv & ","
              ambKeysCsv = ambKeysCsv & CStr(ambKeys(ambI))
            Next
            Call Diag_WriteLine("TX: AMBIGUITY_GATE_SUMMARY count=" & CStr(amb.Count) & " keys=[" & ambKeysCsv & "]")

            ambTotal = amb.Count
            ambShown = ambTotal
            If ambShown > 5 Then ambShown = 5

            For ambI = 0 To ambShown - 1
              ambKey = CStr(ambKeys(ambI))
              If IsObject(amb(ambKey)) Then
                If UCase(TypeName(amb(ambKey))) = "DICTIONARY" Then
                  Dim ambHit, ambDecision, ambInput, ambState, ambCand, ambReason, ambAction
                  Set ambHit = amb(ambKey)
                  ambDecision = "UNKNOWN_DECISION"
                  ambInput = ""
                  ambState = "BLOCKED"
                  ambCand = "(none recorded)"
                  ambReason = ""
                  ambAction = ""
                  If ambHit.Exists("decision_key") Then ambDecision = CStr(ambHit("decision_key"))
                  If ambHit.Exists("input_token") Then ambInput = CStr(ambHit("input_token"))
                  If ambHit.Exists("state") Then ambState = CStr(ambHit("state"))
                  If ambHit.Exists("candidates") Then ambCand = CStr(ambHit("candidates"))
                  If ambHit.Exists("notes") Then ambReason = CStr(ambHit("notes"))
                  If ambHit.Exists("hint") Then ambAction = CStr(ambHit("hint"))
                  Call Diag_WriteLine("TX: AMBIGUITY_DETAIL[" & CStr(ambI + 1) & "] key=" & ambKey & _
                    " decision=" & ambDecision & _
                    " input=[" & Ambiguity_SafeTruncate(ambInput, 80) & "]" & _
                    " state=" & ambState & _
                    " candidates=[" & Ambiguity_SafeTruncate(ambCand, 220) & "]" & _
                    " reason=[" & Ambiguity_SafeTruncate(ambReason, 220) & "]" & _
                    " action=[" & Ambiguity_SafeTruncate(ambAction, 220) & "]")
                Else
                  Call Diag_WriteLine("TX: AMBIGUITY_DETAIL - " & CStr(amb(ambKey)))
                End If
              Else
                Call Diag_WriteLine("TX: AMBIGUITY_DETAIL - " & CStr(amb(ambKey)))
              End If
            Next
 
            If ambTotal > ambShown Then
              Call Diag_WriteLine("TX: AMBIGUITY_DETAIL_TRUNCATED total=" & CStr(ambTotal) & " shown=" & CStr(ambShown))
            End If
          Else
            Call Diag_WriteLine("TX: AMBIGUITY_GATE_SUMMARY count=" & CStr(amb.Count) & " keys=[]")
          End If
          Call Diag_WriteLine("Ambiguity detected. No changes applied.")
          If (PlanValidationErrors Is Nothing) Then
            Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
          End If
          If PlanValidationErrors.Count = 0 Then
            PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
          End If
          Exit Function
        Else
          Call Diag_WriteLine("TX: Ambiguity gate bypassed (allow_ambiguous_apply=True)")
        End If
      End If
    End If
  End If

  Call Diag_WriteLine("TX: TRANSACTION_MODE=" & CStr(TRANSACTION_MODE) & " ApplyPlan.Count=" & CStr(ApplyPlan.Count))

  If ApplyPlan.Count = 0 Then
    PlanValidationErrors.RemoveAll
    PlanValidationErrors("EMPTY_PLAN") = "No tabfields were staged for apply."
    Call Diag_WriteLine("TX: EMPTY_PLAN (no writes staged)")
    Stage_ValidatePlan = False
    Call Diag_WriteLine("TX: EARLY EXIT - EMPTY_PLAN")
    If (PlanValidationErrors Is Nothing) Then
      Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
    End If
    If PlanValidationErrors.Count = 0 Then
      PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
    End If
    Exit Function
  End If

  If PlanValidationErrors.Exists("QUALIFIER_UNRESOLVED") Then
    Stage_ValidatePlan = False
    Call Diag_WriteLine("TX: QUALIFIER_UNRESOLVED - blocking apply")
    Call Diag_WriteLine("TX: EARLY EXIT - QUALIFIER_UNRESOLVED_FINALCHECK")
    If (PlanValidationErrors Is Nothing) Then
      Set PlanValidationErrors = CreateObject("Scripting.Dictionary")
    End If
    If PlanValidationErrors.Count = 0 Then
      PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
    End If
    Exit Function
  End If

  Dim k, v, hasOpen, hasClose
  For Each k In ApplyPlan.Keys
    v = CStr(ApplyPlan(k))

    ' Allow explicit clears (used by output_map clear list)
    If Len(Trim(v)) = 0 Then
      ' OK
    ElseIf Left(UCase(v), 9) = "SMARTSTAT=" Then
      ' OK (control custom property in A)
    Else
      hasOpen  = (InStr(v, "{{") > 0)
      hasClose = (InStr(v, "}}") > 0)

      ' Only enforce pairing if any moustache token is present
      If hasOpen Xor hasClose Then
        Stage_ValidatePlan = False
        PlanValidationErrors(CStr(k)) = "Unbalanced moustache braces in: " & v
      End If
    End If
  Next

  If Not Stage_ValidatePlan Then
    If Not (PlanValidationErrors Is Nothing) Then
      If PlanValidationErrors.Count = 0 Then
        PlanValidationErrors("UNKNOWN_VALIDATE_FAIL") = "Stage_ValidatePlan returned False with no recorded errors."
      End If
    End If

    Dim ek
    For Each ek In PlanValidationErrors.Keys
      Call Diag_WriteLine("TX: VALIDATION ERROR - " & ek & ": " & PlanValidationErrors(ek))
    Next
  Else
    Call Diag_WriteLine("TX: Validation OK")
  End If
End Function

' ==========================================
' v4.0 Stage 5 � Transaction Commit
' ==========================================
Sub Tx_ResetWriteVerifyState()
  G_TRIO_WRITE_ATTEMPTS = 0
  G_TRIO_WRITE_SUCCESS_COUNT = 0
  G_TRIO_WRITE_FAIL_COUNT = 0
  G_TRIO_WRITE_FAILED = False
  G_TRIO_WRITE_LAST_FAIL_TF = ""
End Sub

Sub Stage_CommitTransaction()
  Dim k
  Call Tx_ResetWriteVerifyState()
  Call Diag_WriteLine("TX: Transaction Commit Started - " & ApplyPlan.Count & " fields")
  For Each k In ApplyPlan.Keys
    Call Tx_WriteNow(CStr(k), CStr(ApplyPlan(k)))
    If CBool(G_TRIO_WRITE_FAILED) Then
      Call Diag_WriteLine("TX: Transaction Commit Aborted - Trio write verification failed at tf=" & CStr(G_TRIO_WRITE_LAST_FAIL_TF))
      Exit For
    End If
  Next

  If CBool(G_TRIO_WRITE_FAILED) Then
    If Not (PlanValidationErrors Is Nothing) Then
      PlanValidationErrors("TRIO.CP_WRITE_FAIL") = "Failed to persist tabfield custom property."
    End If
  Else
    Call Diag_WriteLine("TX: Transaction Commit Completed")
    If CLng(G_TRIO_WRITE_SUCCESS_COUNT) > 0 Then
      Diag_Mark_PushToTrio "Category syntax applied to Trio custom properties (writes=" & CStr(G_TRIO_WRITE_SUCCESS_COUNT) & ")"
    End If
  End If
End Sub

' ---------------- Dict helpers ----------------
Function CreateTextDict()
  Dim d : Set d = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  d.CompareMode = 1
  On Error GoTo 0
  Set CreateTextDict = d
End Function

Function NewTextDict() : Set NewTextDict = CreateTextDict() : End Function

Function NormalizeKeyForLookup(k)
  Dim s: s = CStr(k)
  s = Trim(s)
  Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
  s = Replace(s, " ", "_")
  s = LCase(s)
  NormalizeKeyForLookup = s
End Function

Function NormalizeKey(s)
  Dim t, bad, i
  t = LCase(Trim(CStr(s)))
  t = Replace(t, vbTab, " ")
  Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
  bad = "-.':"";\/|()[]{}?"
  For i = 1 To Len(bad): t = Replace(t, Mid(bad, i, 1), ""): Next
  t = Replace(t, " ", "_")
  NormalizeKey = t
End Function

' ---------------- INI ----------------
Function LoadIni(path)
  On Error Resume Next

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(path) Then Set LoadIni = Nothing: Exit Function

  Dim tf: Set tf = fso.OpenTextFile(path, 1, False)

  Dim dict: Set dict = NewTextDict()
  Dim sec, line, eq, k, v
  Dim firstLineHandled: firstLineHandled = False
  sec = ""

  Do Until tf.AtEndOfStream
    ' --- Read the next line; if it's the first line, strip a UTF-8 BOM (U+FEFF) if present ---
    If Not firstLineHandled Then
      line = tf.ReadLine
      If Len(line) > 0 Then
        If AscW(Left(line, 1)) = &HFEFF Then line = Mid(line, 2)
      End If
      firstLineHandled = True
    Else
      line = tf.ReadLine
    End If

    line = Trim(line)

    If Len(line) = 0 Then
      ' skip
    ElseIf Left(line,1) = ";" Or Left(line,1) = "#" Then
      ' skip comments
    ElseIf Left(line,1) = "[" And Right(line,1) = "]" Then
      sec = Mid(line, 2, Len(line)-2)
      If Not dict.Exists(sec) Then dict.Add sec, NewTextDict()
    Else
      eq = InStr(line, "=")
      If eq > 0 And Len(sec) > 0 Then
        k = Trim(Left(line, eq-1))
        v = Trim(Mid(line, eq+1))
        dict(sec)(k) = v
        dict(sec)(NormalizeKeyForLookup(k)) = v
      End If
    End If
  Loop

  tf.Close
  Set LoadIni = dict
  On Error GoTo 0
End Function

Sub LoadIniSectionDictNormalized(ini, sectionName, ByRef rawDict, ByRef normDict)
  Dim k, v
  Set rawDict = NewTextDict()
  Set normDict = NewTextDict()
  If ini Is Nothing Then Exit Sub
  If ini.Exists(sectionName) Then
    Dim sec: Set sec = ini(sectionName)
    For Each k In sec.Keys
      v = sec(k)
      If Not rawDict.Exists(k) Then rawDict.Add k, v
      Dim nk: nk = NormalizeKey(CStr(k))
      If Not normDict.Exists(nk) Then normDict.Add nk, v
    Next
  End If
End Sub

' ---------------- Transforms ----------------
Function EscapeForCharClass(s)
  Dim i, ch, out : out = ""
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    If ch = "\" Or ch = "]" Or ch = "-" Then out = out & "\" & ch Else out = out & ch
  Next
  EscapeForCharClass = out
End Function

Function LoadRegexDict(ini, sectionName)
  On Error Resume Next
  Dim out: Set out = NewTextDict()
  If ini.Exists(sectionName) Then
    Dim sec: Set sec = ini(sectionName)
    Dim k, line, parts
    For Each k In sec.Keys
      line = sec(k)
      parts = Split(line, "=>")
      If UBound(parts) >= 1 Then out(Trim(parts(0))) = Trim(parts(1))
    Next
  End If
  Set LoadRegexDict = out
  On Error GoTo 0
End Function

Function LoadTransforms(ini)
  On Error Resume Next
  Dim t: Set t = NewTextDict()
  If ini.Exists("TRANSFORMS") Then
    Dim sec: Set sec = ini("TRANSFORMS")
    t("to_lower") = (LCase(sec("to_lower"))="true")
    t("trim") = (LCase(sec("trim"))="true")
    t("collapse_spaces") = (LCase(sec("collapse_spaces"))="true")
    t("spaces_to_underscores") = (LCase(sec("spaces_to_underscores"))="true")
    t("normalize_diacritics") = (LCase(sec("normalize_diacritics"))="true")
    If sec.Exists("strip_chars") Then t("strip_chars") = sec("strip_chars") Else t("strip_chars") = ""
  End If
  Set LoadTransforms = t
  On Error GoTo 0
End Function

Function ApplyTransforms(s, transforms, rxTransforms)
  Dim k, prn
  k = CStr(s)
  If transforms.Exists("trim") And transforms("trim") Then k = Trim(k)
  If transforms.Exists("normalize_diacritics") And transforms("normalize_diacritics") Then k = StripDiacritics(k)
  If transforms.Exists("collapse_spaces") And transforms("collapse_spaces") Then Do While InStr(k, "  ") > 0: k = Replace(k, "  ", " "): Loop
  If transforms.Exists("strip_chars") Then
    Dim sc, rx
    sc = CStr(transforms("strip_chars"))
    If Len(sc) > 0 Then
      Set rx = CreateObject("VBScript.RegExp")
      rx.Global = True: rx.IgnoreCase = False
      rx.Pattern = "[" & EscapeForCharClass(sc) & "]"
      k = rx.Replace(k, "")
    End If
  End If
  If transforms.Exists("to_lower") And transforms("to_lower") Then k = LCase(k)
  If transforms.Exists("spaces_to_underscores") And transforms("spaces_to_underscores") Then k = Replace(k, " ", "_")
  For Each prn In rxTransforms.Keys
    k = RegexReplaceAll(k, prn, rxTransforms(prn))
  Next
  ApplyTransforms = k
End Function

Function StripDiacritics(s)
  Dim src, dst, i
  src = "áàäâãåĝçĝéèëêěíìïîľĺńñóòöôõřśšťúùüûýžÝÀÄÂÃÅČÇĎÉÈËÊĚÝÌÝÎĽĹŃÑÓÒÖÔÕŘŚŠŤÚÙÜÛÝŽ"
  dst = "aaaaaaccdeeeeeiiiillnnooooorsstuuuuyzAAAAAACCDEEEEEIIIILLNNOOOOORSSTUUUUYZ"
  For i = 1 To Len(src): s = Replace(s, Mid(src, i, 1), Mid(dst, i, 1)): Next
  StripDiacritics = s
End Function

Function RegexReplaceAll(text, pattern, replacement)
  Dim rx: Set rx = CreateObject("VBScript.RegExp")
  rx.Global = True: rx.IgnoreCase = True
  rx.Pattern = pattern
  replacement = Replace(replacement, "\1", "$1")
  replacement = Replace(replacement, "\2", "$2")
  replacement = Replace(replacement, "\3", "$3")
  RegexReplaceAll = rx.Replace(text, replacement)
End Function

' ---------------- Learn knobs ----------------
Function LoadLearnKnobs(ini)
  Dim d: Set d = NewTextDict()
  d("fuzzy_threshold") = 0.74
  d("fuzzy_threshold_short") = 0.82
  d("phonetic_enable") = False
  d("jaccard_threshold") = 0.68
  ' v4.0 Phase 2 ambiguity knobs
  d("ambiguous_score_delta") = 0.03          ' if bestScore - altScore <= delta => ambiguous
  d("allow_ambiguous_apply") = False         ' broadcast-grade default: block apply
  d("stopwords") = Array("batting","bat","hitting","offensive","rate","percent","percentage","team")
  d("deny_fuzzy") = Array("OBP","OPS","ERA","WHIP","WAR")
  d("prefer_pitcher_tokens") = Array("fb","velo","spin","whiff","csw")

  If ini.Exists("LEARN") Then
    Dim sec: Set sec = ini("LEARN")
    If sec.Exists("fuzzy_threshold") Then d("fuzzy_threshold") = CDbl(sec("fuzzy_threshold"))
    If sec.Exists("fuzzy_threshold_short") Then d("fuzzy_threshold_short") = CDbl(sec("fuzzy_threshold_short"))
    If sec.Exists("phonetic_enable") Then d("phonetic_enable") = (LCase(sec("phonetic_enable"))="true")
    If sec.Exists("jaccard_threshold") Then d("jaccard_threshold") = CDbl(sec("jaccard_threshold"))
    If sec.Exists("ambiguous_score_delta") Then d("ambiguous_score_delta") = CDbl(sec("ambiguous_score_delta"))
    If sec.Exists("allow_ambiguous_apply") Then d("allow_ambiguous_apply") = (LCase(sec("allow_ambiguous_apply"))="true")
    If sec.Exists("stopwords") Then d("stopwords") = ParseCsvList(sec("stopwords"), False)
    If sec.Exists("deny_fuzzy") Then d("deny_fuzzy") = ParseCsvList(sec("deny_fuzzy"), True)
    If sec.Exists("prefer_pitcher_tokens") Then d("prefer_pitcher_tokens") = ParseCsvList(sec("prefer_pitcher_tokens"), False)
  End If
  Set LoadLearnKnobs = d
End Function

Function ParseCsvList(s, toUpper)
  Dim wantUpper
  wantUpper = False
  If VarType(toUpper) = vbBoolean Then
    wantUpper = toUpper
  Else
    wantUpper = (LCase(CStr(toUpper))="true")
  End If
  Dim parts, i : parts = Split(CStr(s), ",")
  For i = 0 To UBound(parts)
    parts(i) = Trim(CStr(parts(i)))
    If wantUpper Then parts(i) = UCase(parts(i))
  Next
  ParseCsvList = parts
End Function

' ---------------- Tokens ----------------
Function StripStopwords(s, stoplist)
  Dim t, i, w
  t = s
  If IsArray(stoplist) Then
    For i = LBound(stoplist) To UBound(stoplist)
      w = Trim(CStr(stoplist(i)))
      If Len(w) > 0 Then t = ReplaceToken(t, w, " ")
    Next
  End If
  Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
  StripStopwords = Trim(t)
End Function

Function ContainsAnyToken(s, tokenList)
  Dim found, i, w, searchSpace, tokenPattern
  found = False
  searchSpace = Replace(" " & s & " ", "_", " ")
  If IsArray(tokenList) Then
    For i = LBound(tokenList) To UBound(tokenList)
      w = Trim(CStr(tokenList(i)))
      If Len(w) > 0 Then
        tokenPattern = Replace(" " & w & " ", "_", " ")
        If InStr(1, searchSpace, tokenPattern, vbTextCompare) > 0 Then
          found = True
          Exit For
        End If
      End If
    Next
  End If
  ContainsAnyToken = found
End Function

Function ReplaceToken(text, token, repl)
  Dim rx: Set rx = CreateObject("VBScript.RegExp")
  rx.Global = True: rx.IgnoreCase = True
  rx.Pattern = "(^|[\s_])" & EscapeRegex(token) & "($|[\s_])"
  ReplaceToken = rx.Replace(text, "$1" & repl & "$2")
End Function

Function EscapeRegex(s)
  Dim t
  t = s
  t = Replace(t, "\", "\\")
  t = Replace(t, ".", "\.")
  t = Replace(t, "+", "\+")
  t = Replace(t, "?", "\?")
  t = Replace(t, "(", "\(")
  t = Replace(t, ")", "\)")
  t = Replace(t, "[", "\[")
  t = Replace(t, "]", "\]")
  t = Replace(t, "{", "\{")
  t = Replace(t, "}", "\}")
  t = Replace(t, "^", "\^")
  t = Replace(t, "$", "\$")
  t = Replace(t, "|", "\|")
  EscapeRegex = t
End Function

' ---------------- Qualifier helpers ----------------
Function NormalizeQualifierPrefixAndKey(rawValue)
  Dim fullValue, prefixType, strippedVal
  Dim re, matches, num
  fullValue = Trim(CStr(rawValue))
  strippedVal = fullValue
  prefixType = ""

  Set re = CreateObject("VBScript.RegExp")
  re.IgnoreCase = True
  re.Global = False

  If UCase(Left(fullValue, 11)) = "THIS SEASON" Then
    prefixType = "season."
    strippedVal = Trim(Mid(fullValue, 12))
  ElseIf UCase(Left(fullValue, 6)) = "SEASON" Then
    prefixType = "season."
    strippedVal = Trim(Mid(fullValue, 7))
  ElseIf UCase(Left(fullValue, 6)) = "CAREER" Then
    prefixType = "career."
    strippedVal = Trim(Mid(fullValue, 7))
  Else
    re.Pattern = "^(LAST|PAST|PREVIOUS)\s+(\d+)\s+GAMES?\s*(.+)?$"
    If re.Test(UCase(fullValue)) Then
      Set matches = re.Execute(UCase(fullValue))
      num = matches(0).SubMatches(1)
      prefixType = "last_game(" & num & ")."
      If Not IsNull(matches(0).SubMatches(2)) Then strippedVal = Trim(matches(0).SubMatches(2)) Else strippedVal = ""
    End If
  End If
  NormalizeQualifierPrefixAndKey = Array(prefixType, strippedVal)
End Function

Function ResolveQualifierSmart(qTxt, qAliasNorm, qNorm, learn, ByRef outFrag, ByRef acceptedBy, ByRef scoreOut)
  Dim raw: raw = Trim(CStr(qTxt))
  If Len(raw) = 0 Then ResolveQualifierSmart = False: Exit Function

  Dim keyN: keyN = NormalizeKey(raw)
  If qAliasNorm.Exists(keyN) Then keyN = NormalizeKey(CStr(qAliasNorm(keyN)))
  If qNorm.Exists(keyN) Then outFrag = CStr(qNorm(keyN)) : acceptedBy="direct" : scoreOut=1 : ResolveQualifierSmart=True : Exit Function

  Dim bestA, altA, sA, sAltA, ambA
  If CBool(DIAG_MODE) Then Call Diag_WriteLine("RESOLVER_CANDIDATE_SOURCE scope=qualifier.alias source=DICT_KEYS order=UNSORTED_ENUM tie_rule=FIRST_SEEN")
  If HeuristicPickWithAlt(keyN, qAliasNorm.Keys, learn, bestA, sA, altA, sAltA, ambA) Then
    If ambA Then
      Call Ambiguity_AddEx("qualifier", "alias", raw, "best=[" & bestA & "](" & ScoreStr(sA) & "); alt=[" & altA & "](" & ScoreStr(sAltA) & ")", "HeuristicPickWithAlt tie")
      acceptedBy = "ambiguous/alias": scoreOut = sA
      ResolveQualifierSmart = False
      Exit Function
    End If

    If qAliasNorm.Exists(bestA) Then
      Dim canon: canon = NormalizeKey(CStr(qAliasNorm(bestA)))
      If qNorm.Exists(canon) Then outFrag = CStr(qNorm(canon)) : acceptedBy="fuzzy/alias" : scoreOut=sA : ResolveQualifierSmart=True : Exit Function
    End If
  End If

  Dim bestC, altC, sC, sAltC, ambC
  If CBool(DIAG_MODE) Then Call Diag_WriteLine("RESOLVER_CANDIDATE_SOURCE scope=qualifier.canon source=DICT_KEYS order=UNSORTED_ENUM tie_rule=FIRST_SEEN")
  If HeuristicPickWithAlt(keyN, qNorm.Keys, learn, bestC, sC, altC, sAltC, ambC) Then
    If ambC Then
      Call Ambiguity_AddEx("qualifier", "canon", raw, "best=[" & bestC & "](" & ScoreStr(sC) & "); alt=[" & altC & "](" & ScoreStr(sAltC) & ")", "HeuristicPickWithAlt tie")
      acceptedBy = "ambiguous/canon": scoreOut = sC
      ResolveQualifierSmart = False
      Exit Function
    End If

    If qNorm.Exists(bestC) Then outFrag = CStr(qNorm(bestC)) : acceptedBy="fuzzy/canon" : scoreOut=sC : ResolveQualifierSmart=True : Exit Function
  End If

  ResolveQualifierSmart = False
End Function

Sub WriteLearnPendingQualifier(learnPath, where, txt, score)
  On Error Resume Next
  Dim line: line = txt & " = (suggest) ??? ; " & where & " last_fuzzy=" & ScoreStr(score)
  AppendIniSectionLine learnPath, "[PENDING]", line
  On Error GoTo 0
End Sub

' ================== Dynamic USAGE resolver (plural + category aware) ==================
' Decides the correct USAGE wildcard measure from the row’s fullPath:
'   - pitch_type(<tok>)      ->  arsenal_<plural(tok)>_percentage
'   - pitch_category(<tok>)  ->  pitch_category_<normalized(tok)>_percentage
'
' Returns "" if no token found.
Function ResolveDynamicUsageMeasure(fullPath)
  On Error Resume Next

  Dim rx, m, tok

  Set rx = CreateObject("VBScript.RegExp")
  rx.Global = False
  rx.IgnoreCase = True

  ' 1) Prefer pitch_type(...)
  rx.Pattern = "pitch_type\(\s*([^)]+?)\s*\)"
  If rx.Test(fullPath) Then
    Set m = rx.Execute(fullPath)
    tok = NormalizePitchToken(m(0).SubMatches(0))
    ResolveDynamicUsageMeasure = "arsenal_" & PluralizePitchType(tok) & "_percentage"
    Exit Function
  End If

  ' 2) Fall back to pitch_category(...)
  rx.Pattern = "pitch_category\(\s*([^)]+?)\s*\)"
  If rx.Test(fullPath) Then
    Set m = rx.Execute(fullPath)
    tok = NormalizePitchToken(m(0).SubMatches(0))
    ResolveDynamicUsageMeasure = PitchCategoryMeasure(tok)
    Exit Function
  End If

  ResolveDynamicUsageMeasure = ""
End Function

' Lowercase + trim + normalize hyphen/space -> underscore
Function NormalizePitchToken(tok)
  Dim t
  t = LCase(Trim(CStr(tok)))
  t = Replace(t, " ", "_")
  t = Replace(t, "-", "_")
  NormalizePitchToken = t
End Function

' Map pitch_type tokens to the plural forms you provided
Function PluralizePitchType(tok)
  Select Case tok
    Case "cutter":            PluralizePitchType = "cutters"
    Case "knuckle_ball","knuckleball": PluralizePitchType = "knuckleballs"
    Case "slider":            PluralizePitchType = "sliders"
    Case "slurve":            PluralizePitchType = "slurves"
    Case "sweeper":           PluralizePitchType = "sweepers"
    Case "sinker":            PluralizePitchType = "sinkers"
    Case "eephus":            PluralizePitchType = "eephuses"
    Case "changeup":          PluralizePitchType = "changeups"
    Case "forkball":          PluralizePitchType = "forkballs"
    Case "splitter":          PluralizePitchType = "splitters"
    Case "fastball","fastballs": PluralizePitchType = "fastballs"
    Case "curveball","curveballs","slow_curve","slowcurve": PluralizePitchType = "curveballs"
    Case "screwball":         PluralizePitchType = "screwballs"
    Case "knuckle_curve","knucklecurve": PluralizePitchType = "knucklecurves"  ' note: no underscore per your list
    ' Common variants you might see; map to best fit:
    Case "4_seam","four_seam": PluralizePitchType = "fastballs"
    Case "2_seam","two_seam":  PluralizePitchType = "fastballs"
    Case Else:                 PluralizePitchType = tok  ' last-resort: just append "s" logic is risky; keep token
  End Select
End Function

' Map pitch_category tokens to the category measures you provided
Function PitchCategoryMeasure(tok)
  Select Case tok
    Case "breaking","breaking_balls","breakingball","breakingballs"
      PitchCategoryMeasure = "pitch_category_breaking_balls_percentage"
    Case "fastball","fastballs"
      PitchCategoryMeasure = "pitch_category_fastballs_percentage"
    Case "offspeed","off_speed","offspeeds"
      PitchCategoryMeasure = "pitch_category_offspeeds_percentage"
    Case Else
      ' If an unexpected category arrives, fallback to a generic arsenal% to avoid empty measure
      PitchCategoryMeasure = "arsenal_percentage"
  End Select
End Function

' Normalizes tokens like "4-seam", "knuckle curve", "changeup" -> "4_seam", "knuckle_curve", "changeup"
Function NormalizePitchKey(tok)
  Dim t
  t = LCase(CStr(tok))
  t = Trim(t)
  t = Replace(t, " ", "_")
  t = Replace(t, "-", "_")
  If t = "knucklecurve" Then t = "knuckle_curve"
  If t = "knuckleball" Then t = "knuckle_ball"
  If t = "slowcurve" Then t = "slow_curve"
  NormalizePitchKey = t
End Function
' ======================================================================

' ---------------- Category resolvers ----------------
' Remove any .pitch_type(...) or .pitch_category(...) segments from a path
Function StripPitchFuncs(ByVal s)
  Dim rx
  Set rx = CreateObject("VBScript.RegExp")
  rx.Global = True
  rx.IgnoreCase = True
  rx.Pattern = "\.pitch_(type|category)\([^\)]*\)"
  s = rx.Replace(s, "")
  Do While InStr(s, "..") > 0
    s = Replace(s, "..", ".")
  Loop
  If Len(s) > 0 And Right(s, 1) = "." Then s = Left(s, Len(s) - 1)
  StripPitchFuncs = s
End Function

Function ResolveCategoryMeasure(inputTxt, isPitcher, catAliasesNorm, catPitchAliasesNorm, catNorm, catPitchNorm)
  Dim x: x = NormalizeKey(CStr(inputTxt))
  If isPitcher Then
    If catPitchAliasesNorm.Exists(x) Then x = NormalizeKey(catPitchAliasesNorm(x))
  Else
    If catAliasesNorm.Exists(x) Then x = NormalizeKey(catAliasesNorm(x))
  End If
  If isPitcher Then
    If catPitchNorm.Exists(x) Then ResolveCategoryMeasure = catPitchNorm(x) : Exit Function
  Else
    If catNorm.Exists(x) Then ResolveCategoryMeasure = catNorm(x) : Exit Function
  End If
  ResolveCategoryMeasure = ""
End Function

Function ResolveCategorySmart(inputKey, preferPitcher, learn, _
        catAlias, catMap, catPitchAlias, catPitchMap, _
        ByRef statOut, ByRef conceptOut, ByRef isPitcher, _
        ByRef usedHeuristic, ByRef acceptedBy, ByRef outScore)

  Dim keyTrim: keyTrim = Trim(inputKey)
  Dim keyUpper: keyUpper = UCase(keyTrim)

  If Len(keyTrim) <= 3 Then
    If InList(keyUpper, learn("deny_fuzzy")) Then
      If DirectCatLookup(keyTrim, catAlias, catMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
      If DirectPitchLookup(keyTrim, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
      ResolveCategorySmart = False: outScore = 0: Exit Function
    End If
  End If

  If preferPitcher Then
    If DirectPitchLookup(keyTrim, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
    If DirectCatLookup(keyTrim, catAlias, catMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
  Else
    If DirectCatLookup(keyTrim, catAlias, catMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
    If DirectPitchLookup(keyTrim, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher) Then ResolveCategorySmart = True: Exit Function
  End If

  Dim collapsedEarly: collapsedEarly = CollapseDoubles(keyTrim)
  If LCase(collapsedEarly) <> LCase(keyTrim) Then
    If preferPitcher Then
      If DirectPitchLookup(collapsedEarly, catPitchAlias, catPitchMap, statOut, conceptOut, isPitcher) Then usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99:ResolveCategorySmart=True:Exit Function
      If DirectCatLookup(collapsedEarly,   catAlias,      catMap,       statOut, conceptOut, isPitcher) Then usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99:ResolveCategorySmart=True:Exit Function
    Else
      If DirectCatLookup(collapsedEarly,   catAlias,      catMap,       statOut, conceptOut, isPitcher) Then usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99:ResolveCategorySmart=True:Exit Function
      If DirectPitchLookup(collapsedEarly, catPitchAlias, catPitchMap,  statOut, conceptOut, isPitcher) Then usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99:ResolveCategorySmart=True:Exit Function
    End If
  End If

  Dim aliasKeys: aliasKeys = MergeKeys(catAlias, catPitchAlias)
  If CBool(DIAG_MODE) Then Call Diag_WriteLine("RESOLVER_CANDIDATE_SOURCE scope=category.alias source=MERGE_KEYS_DICT_ENUM order=UNSORTED_ENUM tie_rule=FIRST_SEEN")

  Dim bestAlias, altAlias, s1, sAlt1, amb1
  If HeuristicPickWithAlt(keyTrim, aliasKeys, learn, bestAlias, s1, altAlias, sAlt1, amb1) Then
    If amb1 Then
      Call Ambiguity_AddEx("category", "alias", keyTrim, "best=[" & bestAlias & "](" & ScoreStr(s1) & "); alt=[" & altAlias & "](" & ScoreStr(sAlt1) & ")", "HeuristicPickWithAlt tie")
      usedHeuristic = True: acceptedBy = "ambiguous/alias": outScore = s1
      ResolveCategorySmart = False
      Exit Function
    End If

    Dim tmpVal
    If catAlias.Exists(bestAlias) Then
      Dim canon: canon = catAlias(bestAlias)
      If TryCanonLookupFlexible(canon, catMap, tmpVal) Then
        statOut = tmpVal : conceptOut = "CATEGORY." & canon : isPitcher = False
        usedHeuristic = True: acceptedBy = "fuzzy/alias": outScore = s1
        ResolveCategorySmart = True: Exit Function
      End If
    End If
    If catPitchAlias.Exists(bestAlias) Then
      canon = catPitchAlias(bestAlias)
      If TryCanonLookupFlexible(canon, catPitchMap, tmpVal) Then
        statOut = tmpVal : conceptOut = "CATEGORY_PITCHER." & canon : isPitcher = True
        usedHeuristic = True: acceptedBy = "fuzzy/alias": outScore = s1
        ResolveCategorySmart = True: Exit Function
      End If
    End If
  End If

  Dim canonKeys: canonKeys = MergeKeys(catMap, catPitchMap)
  If CBool(DIAG_MODE) Then Call Diag_WriteLine("RESOLVER_CANDIDATE_SOURCE scope=category.canon source=MERGE_KEYS_DICT_ENUM order=UNSORTED_ENUM tie_rule=FIRST_SEEN")

  Dim bestCanon, altCanon, s2, sAlt2, amb2
  If HeuristicPickWithAlt(keyTrim, canonKeys, learn, bestCanon, s2, altCanon, sAlt2, amb2) Then
    If amb2 Then
      Call Ambiguity_AddEx("category", "canon", keyTrim, "best=[" & bestCanon & "](" & ScoreStr(s2) & "); alt=[" & altCanon & "](" & ScoreStr(sAlt2) & ")", "HeuristicPickWithAlt tie")
      usedHeuristic = True: acceptedBy = "ambiguous/canon": outScore = s2
      ResolveCategorySmart = False
      Exit Function
    End If

    Dim tmpVal2
    If TryCanonLookupFlexible(bestCanon, catMap, tmpVal2) Then
      statOut = tmpVal2 : conceptOut = "CATEGORY." & bestCanon : isPitcher = False
      usedHeuristic = True: acceptedBy = "fuzzy/canon": outScore = s2
      ResolveCategorySmart = True: Exit Function
    End If
    If TryCanonLookupFlexible(bestCanon, catPitchMap, tmpVal2) Then
      statOut = tmpVal2 : conceptOut = "CATEGORY_PITCHER." & bestCanon : isPitcher = True
      usedHeuristic = True: acceptedBy = "fuzzy/canon": outScore = s2
      ResolveCategorySmart = True: Exit Function
    End If
  End If

  Dim collapsed: collapsed = CollapseDoubles(keyTrim)
  If LCase(collapsed) <> LCase(keyTrim) Then
    Dim outv
    If TryCanonLookupFlexible(collapsed, catMap, outv) Then
      statOut = outv : conceptOut = "CATEGORY." & collapsed : isPitcher=False
      usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99
      ResolveCategorySmart=True:Exit Function
    End If
    If TryCanonLookupFlexible(collapsed, catPitchMap, outv) Then
      statOut = outv : conceptOut = "CATEGORY_PITCHER." & collapsed : isPitcher=True
      usedHeuristic=True:acceptedBy="rescue/collapse":outScore=0.99
      ResolveCategorySmart=True:Exit Function
    End If
  End If

  statOut = "": conceptOut = "": isPitcher = False: usedHeuristic = False
  outScore = Max(s1, s2)
  ResolveCategorySmart = False
End Function

Function TryCanonLookupFlexible(k, dict, ByRef outVal)
  Dim kn: kn = NormalizeKeyForLookup(k)
  If dict.Exists(k) Then outVal = dict(k): TryCanonLookupFlexible = True: Exit Function
  If dict.Exists(kn) Then outVal = dict(kn): TryCanonLookupFlexible = True: Exit Function
  Dim kk
  For Each kk In dict.Keys
    If NormalizeKeyForLookup(CStr(kk)) = kn Then outVal = dict(kk): TryCanonLookupFlexible = True: Exit Function
  Next
  TryCanonLookupFlexible = False
End Function

Function DirectCatLookup(k, catAlias, catMap, ByRef statOut, ByRef conceptOut, ByRef isPitcher)
  Dim kn : kn = NormalizeKeyForLookup(k)
  If catAlias.Exists(k) Then
    Dim canon: canon = catAlias(k)
    If TryCanonLookupFlexible(canon, catMap, statOut) Then conceptOut = "CATEGORY." & canon : isPitcher=False : DirectCatLookup=True : Exit Function
  End If
  If catAlias.Exists(kn) Then
    canon = catAlias(kn)
    If TryCanonLookupFlexible(canon, catMap, statOut) Then conceptOut = "CATEGORY." & canon : isPitcher=False : DirectCatLookup=True : Exit Function
  End If
  If TryCanonLookupFlexible(k,  catMap, statOut) Then conceptOut = "CATEGORY." & k  : isPitcher=False : DirectCatLookup=True : Exit Function
  If TryCanonLookupFlexible(kn, catMap, statOut) Then conceptOut = "CATEGORY." & kn : isPitcher=False : DirectCatLookup=True : Exit Function
  DirectCatLookup = False
End Function

Function DirectPitchLookup(k, catPitchAlias, catPitchMap, ByRef statOut, ByRef conceptOut, ByRef isPitcher)
  Dim kn : kn = NormalizeKeyForLookup(k)
  If catPitchAlias.Exists(k) Then
    Dim canon: canon = catPitchAlias(k)
    If TryCanonLookupFlexible(canon, catPitchMap, statOut) Then conceptOut = "CATEGORY_PITCHER." & canon : isPitcher=True : DirectPitchLookup=True : Exit Function
  End If
  If catPitchAlias.Exists(kn) Then
    canon = catPitchAlias(kn)
    If TryCanonLookupFlexible(canon, catPitchMap, statOut) Then conceptOut = "CATEGORY_PITCHER." & canon : isPitcher=True : DirectPitchLookup=True : Exit Function
  End If
  If TryCanonLookupFlexible(k,  catPitchMap, statOut) Then conceptOut = "CATEGORY_PITCHER." & k  : isPitcher=True : DirectPitchLookup=True : Exit Function
  If TryCanonLookupFlexible(kn, catPitchMap, statOut) Then conceptOut = "CATEGORY_PITCHER." & kn : isPitcher=True : DirectPitchLookup=True : Exit Function
  DirectPitchLookup = False
End Function

Function InList(val, arr)
  Dim i, v
  If IsArray(arr) Then
    For i = LBound(arr) To UBound(arr)
      v = CStr(arr(i))
      If StrComp(CStr(val), v, vbTextCompare) = 0 Then InList = True: Exit Function
    Next
  End If
  InList = False
End Function

Function MergeKeys(d1, d2)
  Dim tmp(), count, k
  count = -1
  If Not IsEmpty(d1) Then If IsObject(d1) Then For Each k In d1.Keys: count = count + 1: ReDim Preserve tmp(count): tmp(count) = CStr(k): Next
  If Not IsEmpty(d2) Then If IsObject(d2) Then For Each k In d2.Keys: count = count + 1: ReDim Preserve tmp(count): tmp(count) = CStr(k): Next
  If count < 0 Then MergeKeys = Array() Else MergeKeys = tmp
End Function

Function CollapseDoubles(s)
  Dim i, ch, prev, out
  out = "": prev = ""
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    If LCase(ch) = LCase(prev) Then
    Else
      out = out & ch
      prev = ch
    End If
  Next
  CollapseDoubles = out
End Function

' --- Fuzzy core ---
Function HeuristicPick(keyTrim, candidateKeys, learn, ByRef bestKey, ByRef scoreOut)
  Dim shortThresh: shortThresh = CDbl(learn("fuzzy_threshold_short"))
  Dim longThresh:  longThresh  = CDbl(learn("fuzzy_threshold"))
  Dim lenKey:      lenKey      = Len(keyTrim)
  Dim threshold, ok, score

  ok = FuzzyResolveAdvanced(keyTrim, candidateKeys, learn, bestKey, score)
  scoreOut = score

  If lenKey <= 4 And Len(bestKey) > 0 Then
    Dim d: d = Lev(LCase(CStr(keyTrim)), LCase(CStr(bestKey)))
    If d = 1 Or IsOneAdjTransposition(LCase(CStr(keyTrim)), LCase(CStr(bestKey))) Then
      HeuristicPick = True
      Exit Function
    End If

    Dim den: den = lenKey
    If den = 0 Then den = 1

    Dim oneEditBound: oneEditBound = 1 - (1 / den)
    ' threshold = Min(oneEditBound, shortThresh) — without IIf
    If oneEditBound > shortThresh Then
      threshold = shortThresh
    Else
      threshold = oneEditBound
    End If
  Else
    threshold = longThresh
  End If

  HeuristicPick = (ok And score >= threshold)
End Function

' ==========================================
' v4.0 Phase 2: Heuristic pick with alternate candidate
' ==========================================
Function HeuristicPickWithAlt(keyTrim, candidateKeys, learn, ByRef bestKey, ByRef bestScore, ByRef altKey, ByRef altScore, ByRef isAmbiguous)
  Dim shortThresh: shortThresh = CDbl(learn("fuzzy_threshold_short"))
  Dim longThresh:  longThresh  = CDbl(learn("fuzzy_threshold"))

  Dim lenKey: lenKey = Len(keyTrim)
  If lenKey < 1 Then lenKey = 1

  Dim ok
  ok = FuzzyResolveTop2Advanced(keyTrim, candidateKeys, learn, bestKey, bestScore, altKey, altScore)

  If (Not ok) Or Len(bestKey) = 0 Then
    HeuristicPickWithAlt = False
    isAmbiguous = False
    Exit Function
  End If

  Dim threshold
  If Len(keyTrim) <= 5 Then
    Dim oneEditBound: oneEditBound = 1 - (1 / lenKey)
    If oneEditBound > shortThresh Then threshold = shortThresh Else threshold = oneEditBound
  Else
    threshold = longThresh
  End If

  Dim passBest: passBest = (bestScore >= threshold)
  Dim passAlt:  passAlt  = (Len(altKey) > 0 And altScore >= threshold)

  isAmbiguous = False
  If passBest And passAlt Then
    Dim delta: delta = CDbl(learn("ambiguous_score_delta"))
    If (bestScore - altScore) <= delta Then isAmbiguous = True
  End If

  HeuristicPickWithAlt = passBest
End Function

Function FuzzyResolveAdvanced(q, candidateKeys, learn, ByRef bestKey, ByRef score)
  Dim shortToken, pref2, pref3, usePhon, sxQ
  Dim bestDist, maxLen
  Dim listA(), listB(), i, k, d
  Dim haveA, haveB
  bestKey = "": score = 0
  bestDist = 9999
  haveA = False: haveB = False

  Dim qn : qn = LCase(CStr(q))
  shortToken = (Len(qn) <= 5)
  pref2 = Left(qn, 2)
  pref3 = Left(qn, 3)

  usePhon = False
  If learn.Exists("phonetic_enable") Then usePhon = CBool(learn("phonetic_enable"))
  If usePhon Then sxQ = Soundex(qn) Else sxQ = ""

  If IsArray(candidateKeys) Then
    For i = LBound(candidateKeys) To UBound(candidateKeys)
      k = LCase(CStr(candidateKeys(i)))
      If shortToken Then
        If (Left(k,3) = pref3) Or (Left(k,2) = pref2) Then AppendString listA, k: haveA = True Else AppendString listB, k: haveB = True
      Else
        AppendString listA, k: haveA = True
      End If
    Next
  Else
    For Each k In candidateKeys
      Dim kk : kk = LCase(CStr(k))
      If shortToken Then
        If (Left(kk,3) = pref3) Or (Left(kk,2) = pref2) Then AppendString listA, kk: haveA = True Else AppendString listB, kk: haveB = True
      Else
        AppendString listA, kk: haveA = True
      End If
    Next
  End If

  If CBool(DIAG_MODE) Then
    Dim rsBefore, rsAfter, rsList
    rsBefore = ""
    Set rsList = CreateObject("System.Collections.ArrayList")

    If haveA Then
      For i = LBound(listA) To UBound(listA)
        If Len(rsBefore) > 0 Then rsBefore = rsBefore & " | "
        rsBefore = rsBefore & CStr(listA(i))
        rsList.Add CStr(listA(i))
      Next
    End If
    If haveB Then
      For i = LBound(listB) To UBound(listB)
        If Len(rsBefore) > 0 Then rsBefore = rsBefore & " | "
        rsBefore = rsBefore & CStr(listB(i))
        rsList.Add CStr(listB(i))
      Next
    End If

    rsAfter = ""
    rsList.Sort
    For i = 0 To rsList.Count - 1
      If Len(rsAfter) > 0 Then rsAfter = rsAfter & " | "
      rsAfter = rsAfter & CStr(rsList(i))
    Next
    Call Diag_WriteLine("RESOLVER_CANDIDATES_BEFORE_SORT q=[" & qn & "] count=" & CStr(rsList.Count) & " list=[" & Ambiguity_SafeTruncate(rsBefore, 320) & "]")
    Call Diag_WriteLine("RESOLVER_CANDIDATES_AFTER_SORT q=[" & qn & "] count=" & CStr(rsList.Count) & " list=[" & Ambiguity_SafeTruncate(rsAfter, 320) & "]")
  End If
  If haveA Then ScanForBest qn, listA, bestKey, bestDist
  If haveB And bestDist > 1 Then ScanForBest qn, listB, bestKey, bestDist

  maxLen = Len(qn): If maxLen < 1 Then maxLen = 1
  score = 1 - (bestDist / maxLen)

  If usePhon Then
    If Soundex(bestKey) = sxQ Then
      score = score + 0.03
      If score > 1 Then score = 1
    End If
  End If
  FuzzyResolveAdvanced = True
End Function

' ==========================================
' v4.0 Phase 2: Top-2 fuzzy resolver (non-breaking addition)
' ==========================================
Function FuzzyResolveTop2Advanced(q, candidateKeys, learn, ByRef bestKey, ByRef bestScore, ByRef altKey, ByRef altScore)
  Dim shortToken, pref2, pref3, usePhon, sxQ
  Dim bestDist, altDist, maxLen
  Dim listA(), listB(), i, k

  bestKey = "": altKey = ""
  bestScore = 0: altScore = 0
  bestDist = 9999: altDist = 9999

  Dim qn : qn = LCase(CStr(q))
  shortToken = (Len(qn) <= 5)
  pref2 = Left(qn, 2)
  pref3 = Left(qn, 3)

  usePhon = False
  If learn.Exists("phonetic_enable") Then usePhon = CBool(learn("phonetic_enable"))
  If usePhon Then sxQ = Soundex(qn) Else sxQ = ""

  Dim haveA: haveA = False
  Dim haveB: haveB = False

  If IsArray(candidateKeys) Then
    For i = LBound(candidateKeys) To UBound(candidateKeys)
      k = LCase(CStr(candidateKeys(i)))
      If shortToken Then
        If (Left(k,3)=pref3) Or (Left(k,2)=pref2) Then AppendString listA, k: haveA=True Else AppendString listB, k: haveB=True
      Else
        AppendString listA, k: haveA=True
      End If
    Next
  Else
    For Each k In candidateKeys
      Dim kk : kk = LCase(CStr(k))
      If shortToken Then
        If (Left(kk,3)=pref3) Or (Left(kk,2)=pref2) Then AppendString listA, kk: haveA=True Else AppendString listB, kk: haveB=True
      Else
        AppendString listA, kk: haveA=True
      End If
    Next
  End If

  If CBool(DIAG_MODE) Then
    Dim rsBefore2, rsAfter2, rsList2
    rsBefore2 = ""
    Set rsList2 = CreateObject("System.Collections.ArrayList")

    If haveA Then
      For i = LBound(listA) To UBound(listA)
        If Len(rsBefore2) > 0 Then rsBefore2 = rsBefore2 & " | "
        rsBefore2 = rsBefore2 & CStr(listA(i))
        rsList2.Add CStr(listA(i))
      Next
    End If
    If haveB Then
      For i = LBound(listB) To UBound(listB)
        If Len(rsBefore2) > 0 Then rsBefore2 = rsBefore2 & " | "
        rsBefore2 = rsBefore2 & CStr(listB(i))
        rsList2.Add CStr(listB(i))
      Next
    End If

    rsAfter2 = ""
    rsList2.Sort
    For i = 0 To rsList2.Count - 1
      If Len(rsAfter2) > 0 Then rsAfter2 = rsAfter2 & " | "
      rsAfter2 = rsAfter2 & CStr(rsList2(i))
    Next
    Call Diag_WriteLine("RESOLVER_CANDIDATES_BEFORE_SORT q=[" & qn & "] count=" & CStr(rsList2.Count) & " list=[" & Ambiguity_SafeTruncate(rsBefore2, 320) & "]")
    Call Diag_WriteLine("RESOLVER_CANDIDATES_AFTER_SORT q=[" & qn & "] count=" & CStr(rsList2.Count) & " list=[" & Ambiguity_SafeTruncate(rsAfter2, 320) & "]")
  End If
  If haveA Then Call ScanForBestTwo(qn, listA, bestKey, bestDist, altKey, altDist)
  If haveB And bestDist > 1 Then Call ScanForBestTwo(qn, listB, bestKey, bestDist, altKey, altDist)

  maxLen = Len(qn): If maxLen < 1 Then maxLen = 1
  bestScore = 1 - (bestDist / maxLen)
  altScore  = 1 - (altDist / maxLen)

  If usePhon Then
    If Len(bestKey) > 0 And Soundex(bestKey) = sxQ Then
      bestScore = bestScore + 0.03: If bestScore > 1 Then bestScore = 1
    End If
    If Len(altKey) > 0 And Soundex(altKey) = sxQ Then
      altScore = altScore + 0.03: If altScore > 1 Then altScore = 1
    End If
  End If

  FuzzyResolveTop2Advanced = True
End Function

Sub ScanForBestTwo(qn, arr, ByRef bestKey, ByRef bestDist, ByRef altKey, ByRef altDist)
  Dim i, k, dist
  For i = LBound(arr) To UBound(arr)
    k = CStr(arr(i))
    dist = Lev(qn, k)

    If CBool(DIAG_MODE) Then
      If Len(bestKey) > 0 And dist = bestDist And k <> bestKey Then
        Call Diag_WriteLine("RESOLVER_TIE q=[" & qn & "] dist=" & CStr(dist) & " incumbent=[" & CStr(bestKey) & "] contender=[" & k & "] tie_rule=FIRST_SEEN")
      End If
      If Len(altKey) > 0 And dist = altDist And k <> altKey Then
        Call Diag_WriteLine("RESOLVER_TIE_ALT q=[" & qn & "] dist=" & CStr(dist) & " incumbent=[" & CStr(altKey) & "] contender=[" & k & "] tie_rule=FIRST_SEEN")
      End If
    End If
    If dist < bestDist Then
      altDist = bestDist
      altKey  = bestKey
      bestDist = dist
      bestKey  = k
    ElseIf dist < altDist And k <> bestKey Then
      altDist = dist
      altKey  = k
    End If

    If Len(qn) <= 5 And bestDist <= 1 Then Exit Sub
  Next
End Sub

Private Sub ScanForBest(qn, arr, ByRef bestKey, ByRef bestDist)
  Dim i, k, dist
  For i = LBound(arr) To UBound(arr)
    k = CStr(arr(i))
    dist = Lev(qn, k)
    If dist < bestDist Then
      bestDist = dist
      bestKey = k
      If Len(qn) <= 5 And bestDist <= 1 Then Exit Sub
    ElseIf dist = bestDist And k <> bestKey Then
      If CBool(DIAG_MODE) Then
        Call Diag_WriteLine("RESOLVER_TIE q=[" & qn & "] dist=" & CStr(dist) & " incumbent=[" & CStr(bestKey) & "] contender=[" & k & "] tie_rule=FIRST_SEEN")
      End If
    End If
  Next
End Sub

Private Sub AppendString(ByRef arr, ByVal s)
  Dim n
  On Error Resume Next
  If IsArray(arr) Then
    n = UBound(arr) + 1
    ReDim Preserve arr(n)
    arr(n) = s
  Else
    ReDim arr(0)
    arr(0) = s
  End If
  On Error GoTo 0
End Sub

Function Lev(a, b)
  Dim la, lb, d(), i, j, cost
  la = Len(a)
  lb = Len(b)
  ReDim d(la, lb)
  For i = 0 To la: d(i,0) = i: Next
  For j = 0 To lb: d(0,j) = j: Next
  For i = 1 To la
    For j = 1 To lb
      cost = 0
      If Mid(a, i, 1) <> Mid(b, j, 1) Then cost = 1
      d(i,j) = Min3(d(i-1,j) + 1, d(i,j-1) + 1, d(i-1,j-1) + cost)
    Next
  Next
  Lev = d(la, lb)
End Function

Function Min3(a,b,c)
  If a < b Then
    If a < c Then Min3 = a Else Min3 = c
  Else
    If b < c Then Min3 = b Else Min3 = c
  End If
End Function

Function IsOneAdjTransposition(a, b)
  If Len(a) = Len(b) And Len(a) >= 2 Then
    Dim i
    For i = 1 To Len(a)-1
      If Mid(a,i,1) <> Mid(b,i,1) Then
        If Mid(a,i,1) = Mid(b,i+1,1) And Mid(a,i+1,1) = Mid(b,i,1) Then
          If Right(a, Len(a)-(i+1)) = Right(b, Len(b)-(i+1)) Then IsOneAdjTransposition = True Else IsOneAdjTransposition = False
          Exit Function
        Else
          Exit For
        End If
      End If
    Next
  End If
  IsOneAdjTransposition = False
End Function

Function Soundex(s)
  s = UCase(s)
  If Len(s) = 0 Then Soundex = "": Exit Function
  Dim first, code, i, ch, d, lastd
  first = Mid(s,1,1)
  code = first
  lastd = ""
  For i = 2 To Len(s)
    ch = Mid(s,i,1)
    d = SoundexDigit(ch)
    If d <> "" And d <> lastd Then code = code & d
    If d <> "" Then lastd = d
    If Len(code) >= 4 Then Exit For
  Next
  If Len(code) < 4 Then code = code & String(4-Len(code),"0")
  Soundex = code
End Function

Function SoundexDigit(ch)
  Select Case ch
    Case "B","F","P","V": SoundexDigit = "1"
    Case "C","G","J","K","Q","S","X","Z": SoundexDigit = "2"
    Case "D","T": SoundexDigit = "3"
    Case "L": SoundexDigit = "4"
    Case "M","N": SoundexDigit = "5"
    Case "R": SoundexDigit = "6"
    Case Else: SoundexDigit = ""
  End Select
End Function

Function ScoreStr(x)
  Dim s: s = FormatNumber(x, 2, -1, 0, -1)
  ScoreStr = Replace(s, ",", "")
End Function

Function Max(a,b)
	If a > b Then
		Max = a
	Else
		Max = b
	End If
End Function

' ---------------- Trio helpers ----------------
Sub TrioSet(tab, val)
  On Error Resume Next
  TrioCmd "page:set_property " & tab & " " & Quote(val)
  On Error GoTo 0
End Sub


' League-aware token post-processor
Function ApplyLeagueNameAdjustments(s, league)
  Dim out: out = CStr(s)
  Select Case UCase(Trim(CStr(league)))
    Case "NHL", "NBA"
      ' Replace preferred_name with first_name for NHL/NBA
      out = Replace(out, "{{info.player.preferred_name}}", "{{info.player.first_name}}")
    Case Else
      ' MLB or unspecified: leave tokens as-is
  End Select
  ApplyLeagueNameAdjustments = out
End Function


        ' Remove redundant ".season.season(<year>)" segments -> ".season(<year>)"
        Function NormalizeSeasonChaining(s)
          Dim rx, out
          out = CStr(s)
          Set rx = CreateObject("VBScript.RegExp")
          rx.Global = True
          rx.IgnoreCase = True
          ' Only collapse when the inner season() uses a 4-digit year (defined season)
          rx.Pattern = "season\.season\(\s*(\d{4})\s*\)"
          out = rx.Replace(out, "season($1)")
          NormalizeSeasonChaining = out
        End Function

Function GetSeasonTrailingSegmentRegex()
  If G_REGEX_SEASON_TRAILING_SEGMENT Is Nothing Then
    Set G_REGEX_SEASON_TRAILING_SEGMENT = CreateObject("VBScript.RegExp")
    G_REGEX_SEASON_TRAILING_SEGMENT.Global = True
    G_REGEX_SEASON_TRAILING_SEGMENT.IgnoreCase = False
    G_REGEX_SEASON_TRAILING_SEGMENT.Pattern = "season\(([0-9]{4})\)\.season\."
  End If
  Set GetSeasonTrailingSegmentRegex = G_REGEX_SEASON_TRAILING_SEGMENT
End Function

    ' Collapse duplicate season chaining and trailing .season after season(<YEAR>)
Function NormalizeSeasonInSyntax(s)
  Dim t: t = CStr(s)

  ' 1) Collapse any ".season.season(" and "season.season(" to a single "season("
  t = Replace(t, ".season.season(", ".season(")
  t = Replace(t, "season.season(", "season(")

  ' 2) Remove extra ".season" immediately after defined year "season(<YEAR>)"
  Dim rx: Set rx = GetSeasonTrailingSegmentRegex()
  t = rx.Replace(t, "season($1).")

  NormalizeSeasonInSyntax = t
End Function

Function Quote(s)
  Quote = Chr(34) & NormalizeSeasonInSyntax(ApplyLeagueNameAdjustments(s, G_SPORT_TAG)) & Chr(34)
End Function

' ==========================================
' v4.0 Transaction Writer Helper (Phase 1)
' ==========================================

Function Diag_TruncateForLog(ByVal s, ByVal maxLen)
  Dim t, m
  t = CStr(s)
  m = CLng(maxLen)
  If m < 0 Then m = 0
  If Len(t) > m Then
    Diag_TruncateForLog = Left(t, m) & "...(len=" & CStr(Len(t)) & ")"
  Else
    Diag_TruncateForLog = t
  End If
End Function

Sub Tx_WriteNow(tfName, value)
  On Error Resume Next

  Dim tf, expected, rc
  Dim setErrNum, setErrDesc
  Dim rb, rbErrNum, rbErrDesc
  Dim verifyWrite, ok, skipNotFound
  Dim isSetNotFound, isReadNotFound

  tf = CStr(tfName)
  expected = NormalizeSeasonInSyntax(ApplyLeagueNameAdjustments(CStr(value), G_SPORT_TAG))

  Err.Clear
  rc = TrioCmd("tabfield:set_custom_property " & tf & " " & Quote(CStr(value)))
  setErrNum = Err.Number
  setErrDesc = CStr(Err.Description)
  Err.Clear
  isSetNotFound = False
  If InStr(1, CStr(rc), "400: Tabfield not found", vbTextCompare) > 0 Then isSetNotFound = True
  If InStr(1, CStr(setErrDesc), "400: Tabfield not found", vbTextCompare) > 0 Then isSetNotFound = True

  G_TRIO_WRITE_ATTEMPTS = CLng(G_TRIO_WRITE_ATTEMPTS) + 1
  verifyWrite = CBool(DIAG_MODE)
  ok = (setErrNum = 0)

  rb = ""
  rbErrNum = 0
  rbErrDesc = ""
  isReadNotFound = False

  If verifyWrite Then
    rb = CStr(TrioCmd("tabfield:get_custom_property " & tf))
    rbErrNum = Err.Number
    rbErrDesc = CStr(Err.Description)
    Err.Clear
    If InStr(1, CStr(rb), "400: Tabfield not found", vbTextCompare) > 0 Then isReadNotFound = True
    If InStr(1, CStr(rbErrDesc), "400: Tabfield not found", vbTextCompare) > 0 Then isReadNotFound = True

    If rbErrNum <> 0 Then ok = False
    If CStr(rb) <> CStr(expected) Then ok = False
  End If

  skipNotFound = False
  If isSetNotFound Or isReadNotFound Then
    If (setErrNum = 0 Or isSetNotFound) And (rbErrNum = 0 Or isReadNotFound) Then
      skipNotFound = True
    End If
  End If

  If skipNotFound Then
    Call Diag_WriteLine( _
      "TRIO_WRITE_SKIP_TABFIELD_NOT_FOUND tf=" & tf & _
      " expected=[" & Diag_TruncateForLog(expected, 180) & "]" & _
      " actual=[" & Diag_TruncateForLog(rb, 180) & "]" & _
      " rc=[" & Diag_TruncateForLog(CStr(rc), 80) & "]" & _
      " setErr=" & CStr(setErrNum) & ":" & Diag_TruncateForLog(setErrDesc, 120) & _
      " readErr=" & CStr(rbErrNum) & ":" & Diag_TruncateForLog(rbErrDesc, 120))
  ElseIf ok Then
    G_TRIO_WRITE_SUCCESS_COUNT = CLng(G_TRIO_WRITE_SUCCESS_COUNT) + 1
  Else
    G_TRIO_WRITE_FAIL_COUNT = CLng(G_TRIO_WRITE_FAIL_COUNT) + 1
    G_TRIO_WRITE_FAILED = True
    G_TRIO_WRITE_LAST_FAIL_TF = tf

    Call Diag_WriteLine( _
      "TRIO_WRITE_VERIFY_FAIL tf=" & tf & _
      " expected=[" & Diag_TruncateForLog(expected, 180) & "]" & _
      " actual=[" & Diag_TruncateForLog(rb, 180) & "]" & _
      " rc=[" & Diag_TruncateForLog(CStr(rc), 80) & "]" & _
      " setErr=" & CStr(setErrNum) & ":" & Diag_TruncateForLog(setErrDesc, 120) & _
      " readErr=" & CStr(rbErrNum) & ":" & Diag_TruncateForLog(rbErrDesc, 120))

    Call Diag_HardFail(PHASE_08_PUSH_TO_TRIO, "TRIO.CP_WRITE_FAIL", "Failed to persist tabfield custom property", "See diag for tfName/rc/readback")
  End If

  On Error GoTo 0
End Sub

Sub Tx_SetCustomProp(tfName, value)
  ' Stages writes when TRANSACTION_MODE=True, otherwise writes immediately.
  If TRANSACTION_MODE Then
    ApplyPlan(CStr(tfName)) = CStr(value)
  Else
    Call Tx_WriteNow(tfName, value)
  End If
End Sub

Sub GuiErr(msg)
  On Error Resume Next
  TrioCmd "gui:error_message " & Quote(msg)
  On Error GoTo 0
End Sub

' ---------------- Logging / FS ----------------
Sub EnsureLogDir(path)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folder: folder = fso.GetParentFolderName(path)
  If Len(folder) > 0 Then If Not fso.FolderExists(folder) Then fso.CreateFolder folder
  On Error GoTo 0
End Sub

' Prevents runaway growth of LOG_FILE (SmartStat_LearnDebug.txt)
Sub Diag_TrimLogFileSize(ByVal path)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        If fso.GetFile(path).Size > DIAG_MAX_FILE_SIZE_BYTES Then
            fso.DeleteFile path, True
        End If
    End If
    On Error GoTo 0
End Sub

Sub LogLine(path, s)
  On Error Resume Next
  Diag_TrimLogFileSize path
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim tf: Set tf = fso.OpenTextFile(path, 8, True)
  tf.WriteLine Now() & " [v4.0.0_beta] " & s
  tf.Close
  On Error GoTo 0
End Sub

Function GetEnv(name)
  On Error Resume Next
  Dim sh: Set sh = CreateObject("WScript.Shell")
  GetEnv = sh.ExpandEnvironmentStrings("%" & name & "%")
  If GetEnv = "%" & name & "%" Then GetEnv = ""
  On Error GoTo 0
End Function

Function ResolveMappingsPath()
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim candidates(5)
  Dim i, cwd, envDir
  envDir = GetEnv("SMARTSTAT_DIR")
  If Len(envDir) > 0 Then
    candidates(0) = fso.BuildPath(envDir, "SmartStat_Mappings.ini")
  Else
    candidates(0) = ""
  End If
  candidates(1) = "SmartStat_Mappings.ini"
  cwd = fso.GetAbsolutePathName(".")
  candidates(2) = fso.BuildPath(cwd, "SmartStat_Mappings.ini")
  candidates(3) = fso.BuildPath(cwd, "SmartStat\SmartStat_Mappings.ini")
  candidates(4) = "E:\EDRIVE\UNIVERSAL\SmartStat\SmartStat_Mappings.ini"
  candidates(5) = "D:\SmartStat\SmartStat_Mappings.ini"
  For i = 0 To UBound(candidates)
    If Len(candidates(i)) > 0 Then
      If fso.FileExists(candidates(i)) Then ResolveMappingsPath = candidates(i): Exit Function
    End If
  Next
  ResolveMappingsPath = ""
  On Error GoTo 0
End Function

' Resolve sport-specific mappings (tries multiple filename conventions)
Function ResolveSportMappingsPath(baseDir, sportTag, wantLearn)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim suffix: suffix = ""
  If wantLearn Then
    suffix = ".learn.ini"
  Else
    suffix = ".ini"
  End If

  Dim cands(3)
  ' Prefer SmartStat_* convention
  cands(0) = fso.BuildPath(baseDir, "SmartStat_Mappings" & sportTag & suffix)        ' SmartStat_MappingsNHL.ini
  cands(1) = fso.BuildPath(baseDir, "Mappings" & sportTag & suffix)                  ' MappingsNHL.ini
  ' Accept in working dir as well
  cands(2) = "SmartStat_Mappings" & sportTag & suffix
  cands(3) = "Mappings" & sportTag & suffix

  Dim i
  For i = 0 To UBound(cands)
    If Len(cands(i)) > 0 Then
      If fso.FileExists(cands(i)) Then ResolveSportMappingsPath = cands(i): Exit Function
    End If
  Next

  ResolveSportMappingsPath = ""
  On Error GoTo 0
End Function


Sub AppendIniSectionLine(path, section, line)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim exists: exists = fso.FileExists(path)
  Dim content: content = ""
  If exists Then
    Dim rf: Set rf = fso.OpenTextFile(path, 1, False)
    content = rf.ReadAll
    rf.Close
  End If

  ' If the exact line already exists in that section, skip
  If InStr(1, content, section, vbTextCompare) > 0 Then
    Dim secPos: secPos = InStr(1, content, section, vbTextCompare)
    Dim nextSec: nextSec = InStr(secPos + Len(section), content, "[")
    Dim secBody
    If nextSec > 0 Then
      secBody = Mid(content, secPos, nextSec - secPos)
    Else
      secBody = Mid(content, secPos)
    End If
    If InStr(1, secBody, vbCrLf & line & vbCrLf, vbTextCompare) > 0 _
       Or Right(secBody, Len(line) + 2) = vbCrLf & line Then
      Exit Sub ' already present
    End If
  End If

  Dim tf: Set tf = fso.OpenTextFile(path, 8, True)
  If InStr(1, content, section, vbTextCompare) = 0 Then tf.WriteLine section
  tf.WriteLine line
  tf.Close
  On Error GoTo 0
End Sub

' ---------------- Small helpers ----------------
Function SingularizeKey(k)
  Dim s: s = LCase(CStr(k))
  If Len(s) <= 3 Then SingularizeKey = k: Exit Function

  ' lefties -> lefty
  If Right(s, 3) = "ies" Then
    SingularizeKey = Left(k, Len(k) - 3) & "y"
    Exit Function
  End If

  ' ches/shes/xes/zes/sses -> drop "es"
  If Right(s, 4) = "ches" Or Right(s, 4) = "shes" Or _
     Right(s, 3) = "xes"  Or Right(s, 3) = "zes"  Or _
     Right(s, 4) = "sses" Then
    SingularizeKey = Left(k, Len(k) - 2)
    Exit Function
  End If

  ' generic trailing s (avoid "ss")
  If Right(s, 1) = "s" And Right(s, 2) <> "ss" Then
    SingularizeKey = Left(k, Len(k) - 1)
    Exit Function
  End If

  SingularizeKey = k
End Function

Function ResolveQualifierChain(rawTxt, qAliasNorm, qNorm, learn, ByRef fragJoined, ByRef leftoversText)
  Dim norm, tokens, i, n
  Dim found, foundLen, foundFrag, j, kk
  Dim cand, canonKey

  norm = NormalizeKey(CStr(rawTxt))           ' e.g., "with_risp_vs_lhp_in_7th_inning_or_later"
  tokens = Split(norm, "_")
  n = UBound(tokens)

  fragJoined = ""                              ' e.g., "risp(yes).vs_pitch_hand(L).innings(7-25)"
  leftoversText = ""                           ' any unmatched tokens (for learn logging)

  ' ------------------------------------------
  ' v4.0 Phase 2: Operator shorthand ambiguity guard
  ' - "VS HP" => ambiguous between VS LHP / VS RHP
  ' - "VS HB" => ambiguous between VS LHB / VS RHB
  ' Broadcast-grade: NEVER guess. Always record ambiguity and hard fail.
  ' ------------------------------------------
  If norm = "vs_hp" Then
    Call Ambiguity_AddEx("qualifier", "canon", CStr(rawTxt), "vs_lhp|vs_rhp", "Operator shorthand VS HP | Hint: type explicit qualifier VS LHP or VS RHP")
    leftoversText = CStr(rawTxt)
    ResolveQualifierChain = False
    Exit Function
  End If

  If norm = "vs_hb" Then
    Call Ambiguity_AddEx("qualifier", "canon", CStr(rawTxt), "vs_lhb|vs_rhb", "Operator shorthand VS HB | Hint: type explicit qualifier VS LHB or VS RHB")
    leftoversText = CStr(rawTxt)
    ResolveQualifierChain = False
    Exit Function
  End If

  ' ------------------------------------------
  ' v4.0 Phase 2: Full-string fuzzy resolve FIRST
  ' Prevents partial matches like "VS HP" resolving to "VS" and ignoring leftovers.
  ' ------------------------------------------
  If Len(norm) > 0 Then
    Dim directFrag, accBy, sOut
    directFrag = "": accBy = "": sOut = 0

    If ResolveQualifierSmart(norm, qAliasNorm, qNorm, learn, directFrag, accBy, sOut) Then
      fragJoined = directFrag
      leftoversText = ""
      ResolveQualifierChain = True
      Exit Function
    Else
      ' If ResolveQualifierSmart flagged ambiguity, fail here so TX gate can block apply.
      If Left(LCase(CStr(accBy)), 9) = "ambiguous" Then
        leftoversText = CStr(rawTxt)
        ResolveQualifierChain = False
        Exit Function
      End If
    End If
  End If

  ' ------------------------------------------
  ' Span matcher (exact). Requires full coverage.
  ' ------------------------------------------
  i = 0
  Do While i <= n
    found = False
    foundLen = 0
    foundFrag = ""

    ' try the longest span starting at i
    For j = n To i Step -1
      ' build cand = tokens[i..j] joined with underscores
      cand = tokens(i)
      For kk = i + 1 To j
        cand = cand & "_" & tokens(kk)
      Next

      ' alias -> canon
      canonKey = cand
      If qAliasNorm.Exists(canonKey) Then canonKey = NormalizeKey(CStr(qAliasNorm(canonKey)))

      ' canon -> fragment
      If qNorm.Exists(canonKey) Then
        found = True
        foundLen = j - i + 1
        foundFrag = CStr(qNorm(canonKey))
        Exit For
      End If

      ' --- singular fallback for single-token candidates (e.g., curveballs -> curveball) ---
      If Not found And i = j Then
        Dim sing, canon2
        sing = NormalizeKey(SingularizeKey(cand))
        If sing <> cand Then
          canon2 = sing
          If qAliasNorm.Exists(canon2) Then canon2 = NormalizeKey(CStr(qAliasNorm(canon2)))
          If qNorm.Exists(canon2) Then
            found = True
            foundLen = 1
            foundFrag = CStr(qNorm(canon2))
            Exit For
          End If
        End If
      End If
    Next

    If found Then
      If Len(fragJoined) > 0 Then fragJoined = fragJoined & "."
      fragJoined = fragJoined & foundFrag
      i = i + foundLen
    Else
      ' no mapping for tokens(i) starting here; accumulate as leftover and advance
      If Len(leftoversText) > 0 Then leftoversText = leftoversText & " " & tokens(i) Else leftoversText = tokens(i)
      i = i + 1
    End If
  Loop

  ' ------------------------------------------
  ' v4.0 Phase 2: FULL COVERAGE REQUIREMENT
  ' If anything is left over, treat as failure.
  ' ------------------------------------------
  If Len(Trim(leftoversText)) > 0 Then
    ResolveQualifierChain = False
  Else
    ResolveQualifierChain = (Len(fragJoined) > 0)
  End If
End Function

Function SuggestQualifierMapping(rawTxt, qAliasNorm, qNorm, learn, _
                                 ByRef suggestIsAlias, ByRef aliasKeyOut, _
                                 ByRef canonKeyOut, ByRef fragOut, ByRef scoreOut)
  Dim keyN: keyN = NormalizeKey(CStr(rawTxt))
  suggestIsAlias = False
  aliasKeyOut = "": canonKeyOut = "": fragOut = "": scoreOut = 0

  ' --- Quick bridge: HOME/ROAD -> AT HOME / ON ROAD (even if no alias rows exist yet) ---
  If keyN = "home" Then
    If qNorm.Exists("at_home") Then
      suggestIsAlias = True
      aliasKeyOut = "at_home"
      canonKeyOut = "at_home"
      fragOut = CStr(qNorm("at_home"))         ' e.g., location(home)
      scoreOut = 0.99
      SuggestQualifierMapping = True
      Exit Function
    End If
  ElseIf keyN = "road" Then
    If qNorm.Exists("on_road") Then
      suggestIsAlias = True
      aliasKeyOut = "on_road"
      canonKeyOut = "on_road"
      fragOut = CStr(qNorm("on_road"))         ' e.g., location(away)
      scoreOut = 0.99
      SuggestQualifierMapping = True
      Exit Function
    End If
  End If

  ' --- 1) Token containment against ALIAS keys (alias->canon mapping) ---
  Dim k, toks, ti, tok
  For Each k In qAliasNorm.Keys   ' normalized alias key, e.g., "at_home"
    toks = Split(LCase(CStr(k)), "_")
    For ti = LBound(toks) To UBound(toks)
      tok = toks(ti)
      If tok = LCase(keyN) Then
        aliasKeyOut = CStr(k)
        canonKeyOut = NormalizeKey(CStr(qAliasNorm(k)))     ' normalized canonical key
        If qNorm.Exists(canonKeyOut) Then fragOut = CStr(qNorm(canonKeyOut)) Else fragOut = ""
        scoreOut = 0.95
        suggestIsAlias = True
        SuggestQualifierMapping = True
        Exit Function
      End If
    Next
  Next

  ' --- 2) Token containment against CANONICAL keys (e.g., keyN "home" hits "at_home") ---
  For Each k In qNorm.Keys        ' canonical key, e.g., "at_home"
    toks = Split(LCase(CStr(k)), "_")
    For ti = LBound(toks) To UBound(toks)
      tok = toks(ti)
      If tok = LCase(keyN) Then
        canonKeyOut = CStr(k)
        fragOut = CStr(qNorm(k))
        scoreOut = 0.90
        suggestIsAlias = False
        SuggestQualifierMapping = True
        Exit Function
      End If
    Next
  Next

  ' --- 3) Fuzzy among ALIAS keys ---
  Dim bestA, sA
  If HeuristicPick(keyN, qAliasNorm.Keys, learn, bestA, sA) Then
    aliasKeyOut = bestA
    canonKeyOut = NormalizeKey(CStr(qAliasNorm(bestA)))
    If qNorm.Exists(canonKeyOut) Then fragOut = CStr(qNorm(canonKeyOut)) Else fragOut = ""
    scoreOut = sA
    suggestIsAlias = True
    SuggestQualifierMapping = True
    Exit Function
  End If

  ' --- 4) Fuzzy among CANONICAL keys ---
  Dim bestC, sC
  If HeuristicPick(keyN, qNorm.Keys, learn, bestC, sC) Then
    canonKeyOut = bestC
    fragOut = CStr(qNorm(bestC))
    scoreOut = sC
    suggestIsAlias = False
    SuggestQualifierMapping = True
    Exit Function
  End If

  SuggestQualifierMapping = False
End Function

Sub WritePendingQualifierSuggestion(learnPath, rawTxt, suggestIsAlias, aliasKey, canonKey, score, whereTag)
  On Error Resume Next
  If suggestIsAlias Then
    ' Suggest an alias mapping to add under [QUALIFIER_TO_FILTER_ALIASES]
    AppendIniSectionLine learnPath, "[PENDING_QUALIFIER_ALIASES]", _
      rawTxt & " = " & aliasKey & " ; suggest alias " & whereTag & " score=" & ScoreStr(score)
  Else
    ' Suggest a direct canonical key to add under [QUALIFIER_TO_FILTER]
    AppendIniSectionLine learnPath, "[PENDING_QUALIFIER]", _
      rawTxt & " = " & canonKey & " ; suggest canon " & whereTag & " score=" & ScoreStr(score)
  End If
  On Error GoTo 0
End Sub

Function TryAdjacentSwapResolve(inputTxt, isPitcher, catAliasNorm, catPAliasNorm, catNorm, catPNorm)
  Dim s, i, a, b, swapped, m
  s = CStr(inputTxt)
  TryAdjacentSwapResolve = ""
  If Len(s) < 2 Then Exit Function

  For i = 1 To Len(s) - 1
    ' swap chars at i and i+1
    a = Mid(s, i, 1)
    b = Mid(s, i+1, 1)
    swapped = Left(s, i-1) & b & a & Mid(s, i+2)

    m = ResolveCategoryMeasure(swapped, isPitcher, catAliasNorm, catPAliasNorm, catNorm, catPNorm)
    If Len(m) > 0 Then
      TryAdjacentSwapResolve = m
      Exit Function
    End If
  Next
End Function

Function SafeGet(d, key, def)
  If d Is Nothing Then SafeGet = def : Exit Function
  If d.Exists(key) Then SafeGet = CStr(d(key)) Else SafeGet = def
End Function

Function NormalizeTemplateNameForKeys(s)
  Dim t, i, ch
  t = UCase(CStr(s))

  ' strip path prefixes
  If InStr(t, "/") > 0 Then t = Split(t, "/")(UBound(Split(t, "/")))
  Dim bslash: bslash = "\"
  If InStr(t, bslash) > 0 Then
    Dim parts: parts = Split(t, bslash)
    t = parts(UBound(parts))
  End If

  ' strip common extensions
  If Right(t, 4) = ".WIZ" Then t = Left(t, Len(t)-4)
  If Right(t, 4) = ".CFX" Then t = Left(t, Len(t)-4)
  If Right(t, 3) = ".CF"  Then t = Left(t, Len(t)-3)

  ' squash non-alnum to underscore
  Dim out: out = ""
  For i = 1 To Len(t)
    ch = Mid(t, i, 1)
    If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
      out = out & ch
    Else
      out = out & "_"
    End If
  Next
  NormalizeTemplateNameForKeys = out
End Function

Function PushArray(arr, val)
  Dim i, tmp()
  If IsArray(arr) Then
    ReDim tmp(UBound(arr)+1)
    For i = LBound(arr) To UBound(arr)
      tmp(i) = arr(i)
    Next
    tmp(UBound(arr)+1) = val
    PushArray = tmp
  Else
    PushArray = Array(val)
  End If
End Function

Function LoadTemplateSectionConfig(srcDir, tmplNameKey, ByRef tsec, ByRef qualTab, ByRef catTabs, ByRef outItems, ByRef filterTabs, ByRef haveFilters)
  Dim tmplIniPath: tmplIniPath = srcDir & "SmartStat_TemplateConfig.ini"
  Dim variants: Set variants = LoadTemplateVariants(tmplIniPath, "TEMPLATE:" & tmplNameKey)
  If variants Is Nothing Or variants.Count = 0 Then LoadTemplateSectionConfig = False: Exit Function

  Dim cfgHint: cfgHint = UCase(Trim(TrioCmd("page:get_property A")))
  If Len(cfgHint) = 0 Then cfgHint = UCase(Trim(TrioCmd("tabfield:get_custom_property A")))
  Dim forcedAlt: forcedAlt = (InStr(cfgHint, "CONFIG=ALT") > 0)

  Dim bestIdx: bestIdx = -1
  Dim bestScore: bestScore = -9999
  Dim i, sc

  Dim pageTabs: pageTabs = Split(TrioCmd("page:get_tabfield_names"))
  For i = 0 To variants.Count - 1
    If forcedAlt Then
      If UCase(SafeGet(variants(i), "config_id", "DEFAULT")) = "ALT" Then
        bestIdx = i: Exit For
      End If
    Else
      sc = ScoreVariant(variants(i), pageTabs)
      If sc > bestScore Then bestScore = sc: bestIdx = i
      If sc = bestScore Then
        If UCase(SafeGet(variants(i), "config_id", "DEFAULT")) = "DEFAULT" Then bestIdx = i
      End If
    End If
  Next

  If bestIdx < 0 Then LoadTemplateSectionConfig = False: Exit Function

  Set tsec = variants(bestIdx)
  qualTab = SafeGet(tsec, "qualifier", "none")

  Dim filterTabsCsv, catTabsCsv, outMapCsv
  filterTabsCsv = SafeGet(tsec, "filter_tabfields", "")
  catTabsCsv    = SafeGet(tsec, "category_tabfields", "")
  outMapCsv     = SafeGet(tsec, "output_map", "")

  If Len(catTabsCsv) = 0 Then LoadTemplateSectionConfig = False: Exit Function

  catTabs  = Split(catTabsCsv, ",")
  If Len(outMapCsv) > 0 Then
    outItems = Split(outMapCsv, ",")
  Else
    outItems = Array()
  End If

  Dim filterTabsNorm: filterTabsNorm = LCase(Trim(filterTabsCsv))
  haveFilters = False
  If Len(filterTabsCsv) > 0 And filterTabsNorm <> "none" Then
    filterTabs = Split(filterTabsCsv, ",")
    haveFilters = True
  Else
    filterTabs = Array()
  End If

  LoadTemplateSectionConfig = True
End Function

Private Function LoadTemplateVariants(path, targetSectionName)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(path) Then Set LoadTemplateVariants = Nothing: Exit Function

  Dim tf: Set tf = fso.OpenTextFile(path, 1, False)
  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  Dim curName, curDict, line, secHeader, eq, k, v

  curName = ""
  Set curDict = Nothing

  Do Until tf.AtEndOfStream
    line = Trim(tf.ReadLine)
    If Len(line) = 0 Then
    ElseIf Left(line, 1) = ";" Or Left(line, 1) = "#" Then
    ElseIf Left(line,1) = "[" And Right(line,1) = "]" Then
      secHeader = Mid(line, 2, Len(line)-2)
      If UCase(secHeader) = UCase(targetSectionName) Then
        If Not curDict Is Nothing Then list.Add curDict
        Set curDict = NewTextDict()
        curName = secHeader
      Else
        If Not curDict Is Nothing Then list.Add curDict: Set curDict = Nothing
        curName = secHeader
      End If
    Else
      If Not curDict Is Nothing Then
        eq = InStr(line, "=")
        If eq > 0 Then
          k = Trim(Left(line, eq-1))
          v = Trim(Mid(line, eq+1))
          curDict(k) = v
        End If
      End If
    End If
  Loop
  tf.Close
  If Not curDict Is Nothing Then list.Add curDict

  Set LoadTemplateVariants = list
  On Error GoTo 0
End Function

Private Function ScoreVariant(secDict, pageTabs)
  Dim s: s = 0
  Dim q: q = UCase(SafeGet(secDict, "config_id", "DEFAULT"))
  If q = "ALT" Then s = s + 0

  Dim qTab: qTab = SafeGet(secDict, "qualifier", "none")
  If LCase(qTab) <> "none" Then If TabExists(pageTabs, qTab) Then s = s + 2

  Dim cats, i
  cats = Split(SafeGet(secDict, "category_tabfields", ""), ",")
  For i = LBound(cats) To UBound(cats)
    If TabExists(pageTabs, cats(i)) Then s = s + 1
  Next

  Dim filts
  filts = Split(SafeGet(secDict, "filter_tabfields", ""), ",")
  If UBound(filts) >= LBound(filts) Then
    If LCase(Trim(SafeGet(secDict, "filter_tabfields", ""))) <> "none" Then
      For i = LBound(filts) To UBound(filts)
        If TabExists(pageTabs, filts(i)) Then s = s + 1
      Next
    End If
  End If

  ScoreVariant = s
End Function

Private Function TabExists(pageTabs, nameTxt)
  Dim i, n: n = UCase(Trim(CStr(nameTxt)))
  For i = LBound(pageTabs) To UBound(pageTabs)
    If UCase(CStr(pageTabs(i))) = n Then TabExists = True: Exit Function
  Next
  TabExists = False
End Function


Sub LoadCategoryAndQualifierMappings(mappingsIni, ByRef catRaw, ByRef catNorm, ByRef catPRaw, ByRef catPNorm, ByRef catAliasRaw, ByRef catAliasNorm, ByRef catPAliasRaw, ByRef catPAliasNorm, ByRef qRaw, ByRef qNorm, ByRef qAliasRaw, ByRef qAliasNorm)
  LoadIniSectionDictNormalized mappingsIni, "CATEGORY_TO_MEASURE", catRaw, catNorm
  LoadIniSectionDictNormalized mappingsIni, "CATEGORY_TO_MEASURE_PITCHER", catPRaw, catPNorm
  LoadIniSectionDictNormalized mappingsIni, "CATEGORY_TO_MEASURE_ALIASES", catAliasRaw, catAliasNorm
  LoadIniSectionDictNormalized mappingsIni, "CATEGORY_TO_MEASURE_PITCHER_ALIASES", catPAliasRaw, catPAliasNorm

  LoadIniSectionDictNormalized mappingsIni, "QUALIFIER_TO_FILTER", qRaw, qNorm
  LoadIniSectionDictNormalized mappingsIni, "QUALIFIER_TO_FILTER_ALIASES", qAliasRaw, qAliasNorm
End Sub

Sub DetermineEntityContext(ByRef entityCtx, ByRef entityType, ByRef playerSubtype)
  Dim aFlag, aPage
  aFlag = UCase(TrioCmd("tabfield:get_custom_property A"))
  aPage = UCase(TrioCmd("page:get_property A"))

  If InStr(aFlag, "SMARTSTAT=") = 0 And Len(aPage) = 0 Then
    Call Tx_SetCustomProp("A", "SMARTSTAT=PLAYER")
    entityCtx = "player": entityType = "PLAYER": playerSubtype = ""
  Else
    If Len(aPage) > 0 Then
      Dim entityBlock, subparts
      entityBlock = aPage
      If InStr(entityBlock, "-") > 0 Then
        subparts = Split(entityBlock, "-")
        entityType = UCase(Trim(subparts(0)))
        playerSubtype = UCase(Trim(subparts(1)))
      Else
        entityType = UCase(entityBlock)
        playerSubtype = ""
      End If
      entityCtx = LCase(entityType)
      Call Tx_SetCustomProp("A", "SMARTSTAT=" & entityType)
    Else
      If InStr(aFlag, "SMARTSTAT=TEAM") > 0 Then
        entityCtx = "team"
      ElseIf InStr(aFlag, "SMARTSTAT=LEAGUE") > 0 Then
        entityCtx = "league"
      Else
        entityCtx = "player"
        entityType = UCase(entityCtx): playerSubtype = ""
      End If
    End If
  End If
End Sub

Function ProcessQualifier(qualTab, qAliasNorm, qNorm, learn, ByRef qPrefix, ByRef qRemFrag)
  Dim qualTxt, qRemainder, chainOk
  Dim acceptedBy, scoreOut

  Call EnsureAmbiguityContext()
  qualTxt = Trim(CStr(TrioCmd("page:get_property " & CStr(qualTab))))

  ' -----------------------------
  ' Broadcast-grade rule:
  ' - If operator typed NOTHING -> safe default "season"
  ' - If operator typed SOMETHING and we cannot fully resolve -> FAIL (no silent season default)
  ' -----------------------------
  If Len(qualTxt) = 0 Then
    qPrefix = "season"
    qRemFrag = ""
    ProcessQualifier = True
    Exit Function
  End If

  qPrefix = ""
  qRemFrag = ""
  qRemainder = ""

  ' Resolve qualifier chain (full coverage expected; leftovers indicate failure)
  chainOk = ResolveQualifierChain(qualTxt, qAliasNorm, qNorm, learn, qRemFrag, qRemainder)

  If Not chainOk Then
    ' If ambiguity was recorded, Stage_ValidatePlan() will block apply.
    ' Otherwise, treat as unresolved qualifier -> hard fail.
    If Not (CompilerContext Is Nothing) Then
      If CompilerContext.Exists("ambiguous") Then
        If CompilerContext("ambiguous").Count > 0 Then
          Call Ambiguity_AddDetailed("QUALIFIER", qualTxt, CompilerContext("ambiguous").Keys, "BLOCKED", "Multiple qualifier candidates remained above threshold.", "Set " & CStr(qualTab) & " to one exact qualifier token (example: VS_CHANGEUP), or update learn alias mapping.")
          Call Diag_WriteLine("QUALIFIER: ambiguous/unresolved input=[" & qualTxt & "] (ambigCount=" & CStr(CompilerContext("ambiguous").Count) & ")")
          Call Diag_WriteAmbiguitySummary()
          Call Diag_WriteLine("TX: EARLY EXIT - AMBIGUOUS_GATE")
          ProcessQualifier = False
          Exit Function
        End If
      End If
    End If

    PlanValidationErrors.RemoveAll
    PlanValidationErrors("QUALIFIER_UNRESOLVED") = "Qualifier text could not be resolved: [" & qualTxt & "] leftovers=[" & Trim(CStr(qRemainder)) & "]"
    Call Diag_WriteLine("QUALIFIER: UNRESOLVED input=[" & qualTxt & "] leftovers=[" & Trim(CStr(qRemainder)) & "]")
    Call Ambiguity_AddDetailed("QUALIFIER", qualTxt, "", "UNRESOLVED", "Qualifier chain did not fully resolve (leftover tokens remained).", "Use an exact qualifier key in " & CStr(qualTab) & " or add a learn alias for this phrase.")
    Call Diag_WriteAmbiguitySummary()
    Call Diag_WriteLine("TX: EARLY EXIT - QUALIFIER_UNRESOLVED_CHAIN")
    ProcessQualifier = False
    Exit Function
  End If

  ' Determine prefix for qualifier display/labeling (existing behavior)
  ' If your original logic computed qPrefix elsewhere, keep that computation below.
  ' Minimal safe default: use "season" when resolved fragment starts with season(...)
  qPrefix = "season"

  ' v4.0 Phase 3: normalize qualifier parts to prevent duplicate prefixes like "season.season..."
  Call NormalizeQualifierParts(qPrefix, qRemFrag)

  ProcessQualifier = True
End Function

' ------------------------------------------
' v4.0 Phase 3: Qualifier normalization
' Prevents duplicate prefixes when qPrefix and qRemFrag both encode "season"
' Example bad:  qPrefix="season", qRemFrag="season.vsLHP" -> season.season.vsLHP
' Example good: qPrefix="season", qRemFrag="vsLHP"
' ------------------------------------------
Sub NormalizeQualifierParts(ByRef qPrefix, ByRef qRemFrag)
  Dim p, r
  p = LCase(Trim(CStr(qPrefix)))
  r = Trim(CStr(qRemFrag))

  If Len(p) = 0 Then
    ' If remainder begins with season., promote prefix and strip remainder
    If LCase(Left(r, 7)) = "season." Then
      qPrefix = "season"
      qRemFrag = Mid(r, 8)
    End If
    Exit Sub
  End If

  ' Strip duplicate prefix from remainder (case-insensitive): "season." + remainder
  If Len(r) > 0 Then
    If LCase(Left(r, Len(p) + 1)) = p & "." Then
      qRemFrag = Mid(r, Len(p) + 2)
      Exit Sub
    End If
  End If

  ' If remainder equals prefix exactly, clear remainder
  If Len(r) > 0 Then
    If LCase(r) = p Then qRemFrag = ""
  End If
End Sub

Function ArrayHasElements(arr)
  On Error Resume Next
  Dim lb, ub
  lb = LBound(arr)
  ub = UBound(arr)
  If Err.Number <> 0 Then
    Err.Clear
    ArrayHasElements = False
  Else
    ArrayHasElements = (ub >= lb)
  End If
  On Error GoTo 0
End Function

Function NormalizeTabfieldToken(token)
  NormalizeTabfieldToken = UCase(Trim(CStr(token)))
End Function

Function IsDigits4(s)
  Dim i, ch
  IsDigits4 = False
  If Len(s) <> 4 Then Exit Function
  For i = 1 To 4
    ch = Mid(s, i, 1)
    If ch < "0" Or ch > "9" Then Exit Function
  Next
  IsDigits4 = True
End Function

Function ParseTabfieldCode(tabName, ByRef prefixOut, ByRef numOut)
  Dim t, d
  ParseTabfieldCode = False
  prefixOut = ""
  numOut = 0

  t = NormalizeTabfieldToken(tabName)
  If Len(t) <> 5 Then Exit Function

  prefixOut = Left(t, 1)
  d = Mid(t, 2, 4)
  If Not IsDigits4(d) Then Exit Function

  numOut = CInt(d)
  ParseTabfieldCode = True
End Function

Function OutputPrefixPriority(prefixTxt)
  Dim p
  p = UCase(Trim(CStr(prefixTxt)))
  If p = "A" Or p = "E" Then OutputPrefixPriority = 99: Exit Function
  If p >= "H" And p <= "Y" Then OutputPrefixPriority = 0: Exit Function
  If p = "B" Or p = "C" Then OutputPrefixPriority = 2: Exit Function
  If p >= "A" And p <= "Z" Then OutputPrefixPriority = 1: Exit Function
  OutputPrefixPriority = 50
End Function

Function CanInferFromCategoryPrefix(prefixTxt)
  Dim pr
  pr = OutputPrefixPriority(prefixTxt)
  CanInferFromCategoryPrefix = (pr < 2)
End Function

Function TryParseOutputMapItem(itm, ByRef tgtTab, ByRef colIdx, ByRef rowIdx)
  On Error Resume Next
  Dim raw, parts

  TryParseOutputMapItem = False
  tgtTab = ""
  colIdx = 0
  rowIdx = 0

  raw = Trim(CStr(itm))
  If Len(raw) = 0 Then Exit Function

  parts = Split(raw, ":")
  If UBound(parts) < 2 Then Exit Function

  tgtTab = NormalizeTabfieldToken(parts(0))
  If Len(tgtTab) = 0 Then Exit Function
  If Not IsNumeric(Trim(CStr(parts(1)))) Then Exit Function
  If Not IsNumeric(Trim(CStr(parts(2)))) Then Exit Function

  colIdx = CInt(parts(1))
  rowIdx = CInt(parts(2))
  If colIdx <= 0 Or rowIdx <= 0 Then Exit Function

  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If

  TryParseOutputMapItem = True
  On Error GoTo 0
End Function

Function ArrayListToArray(listObj)
  Dim arr(), i
  If listObj Is Nothing Then ArrayListToArray = Array(): Exit Function
  If listObj.Count = 0 Then ArrayListToArray = Array(): Exit Function

  ReDim arr(listObj.Count - 1)
  For i = 0 To listObj.Count - 1
    arr(i) = CStr(listObj(i))
  Next
  ArrayListToArray = arr
End Function

Function CollectOutputCandidatesForGroup(pageTabs, anchorPrefix, anchorGroup, rowCount, catTab, allowCategorySelf)
  Dim prim, sec, seen
  Dim i, t, pfx, numVal, suffix
  Set prim = CreateObject("System.Collections.ArrayList")
  Set sec  = CreateObject("System.Collections.ArrayList")
  Set seen = NewTextDict()

  For i = LBound(pageTabs) To UBound(pageTabs)
    t = NormalizeTabfieldToken(pageTabs(i))
    If Len(t) > 0 Then
      If ParseTabfieldCode(t, pfx, numVal) Then
        If pfx = anchorPrefix And (numVal \ 100) = anchorGroup Then
          suffix = (numVal Mod 100)
          If suffix <> 0 Then
            If allowCategorySelf Or t <> UCase(catTab) Then
              If Not seen.Exists(t) Then
                seen(t) = True
                If rowCount > 1 Then
                  If (suffix Mod 10) = 0 Then
                    prim.Add t
                  Else
                    sec.Add t
                  End If
                Else
                  prim.Add t
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next

  prim.Sort
  sec.Sort

  Dim merged: Set merged = CreateObject("System.Collections.ArrayList")
  For i = 0 To prim.Count - 1
    merged.Add CStr(prim(i))
  Next
  For i = 0 To sec.Count - 1
    merged.Add CStr(sec(i))
  Next

  CollectOutputCandidatesForGroup = ArrayListToArray(merged)
End Function

Function FirstUnusedCandidate(candidates, usedTabs)
  Dim i, c
  FirstUnusedCandidate = ""
  If Not ArrayHasElements(candidates) Then Exit Function

  For i = LBound(candidates) To UBound(candidates)
    c = UCase(Trim(CStr(candidates(i))))
    If Len(c) > 0 Then
      If Not usedTabs.Exists(c) Then
        FirstUnusedCandidate = c
        Exit Function
      End If
    End If
  Next
End Function

Function GetMappedTabForColumn(mappedByKey, colIdx, rowCount, ByRef hasExplicitAnchor)
  Dim r, key
  Dim bestRow, bestTab
  Dim k, parts, cVal, rVal

  hasExplicitAnchor = False
  GetMappedTabForColumn = ""

  For r = 1 To rowCount
    key = CStr(colIdx) & ":" & CStr(r)
    If mappedByKey.Exists(key) Then
      hasExplicitAnchor = True
      GetMappedTabForColumn = NormalizeTabfieldToken(mappedByKey(key))
      Exit Function
    End If
  Next

  bestRow = 2147483647
  bestTab = ""
  For Each k In mappedByKey.Keys
    parts = Split(CStr(k), ":")
    If UBound(parts) >= 1 Then
      If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
        cVal = CInt(parts(0))
        rVal = CInt(parts(1))
        If cVal = colIdx Then
          If rVal < bestRow Then
            bestRow = rVal
            bestTab = NormalizeTabfieldToken(mappedByKey(k))
          End If
        End If
      End If
    End If
  Next

  If bestRow <> 2147483647 Then
    hasExplicitAnchor = True
    GetMappedTabForColumn = bestTab
  End If
End Function

Function BuildEffectiveOutputMap(catTabs, outItems, rowCount, ByRef inferredCount)
  Dim mappedByKey, existingItems, inferredItems
  Dim i, tgtTab, colIdx, rowIdx, key
  Set mappedByKey = NewTextDict()
  Set existingItems = CreateObject("System.Collections.ArrayList")
  Set inferredItems = CreateObject("System.Collections.ArrayList")
  inferredCount = 0

  If ArrayHasElements(outItems) Then
    For i = LBound(outItems) To UBound(outItems)
      If TryParseOutputMapItem(outItems(i), tgtTab, colIdx, rowIdx) Then
        existingItems.Add tgtTab & ":" & CStr(colIdx) & ":" & CStr(rowIdx)
        key = CStr(colIdx) & ":" & CStr(rowIdx)
        If Not mappedByKey.Exists(key) Then mappedByKey(key) = tgtTab
      End If
    Next
  End If

  Dim pageTabs: pageTabs = Split(TrioCmd("page:get_tabfield_names"))
  Dim catCount
  catCount = 0
  If ArrayHasElements(catTabs) Then catCount = UBound(catTabs) - LBound(catTabs) + 1

  Dim c, catTab, anchorTab, hasExplicitAnchor
  Dim anchorPrefix, anchorNum, anchorGroup
  Dim catPrefix, catNum
  Dim candidates, usedTabs
  Dim r, pick

  For c = 0 To catCount - 1
    colIdx = c + 1
    catTab = NormalizeTabfieldToken(catTabs(LBound(catTabs) + c))
    If Len(catTab) > 0 Then
      hasExplicitAnchor = False
      anchorTab = GetMappedTabForColumn(mappedByKey, colIdx, rowCount, hasExplicitAnchor)

      If Len(anchorTab) = 0 Then
        If ParseTabfieldCode(catTab, catPrefix, catNum) Then
          If CanInferFromCategoryPrefix(catPrefix) Then
            anchorTab = catTab
          End If
        End If
      End If

      If Len(anchorTab) > 0 Then
        If ParseTabfieldCode(anchorTab, anchorPrefix, anchorNum) Then
          anchorGroup = anchorNum \ 100
          candidates = CollectOutputCandidatesForGroup(pageTabs, anchorPrefix, anchorGroup, rowCount, catTab, hasExplicitAnchor)
          Set usedTabs = NewTextDict()

          For r = 1 To rowCount
            key = CStr(colIdx) & ":" & CStr(r)
            If mappedByKey.Exists(key) Then usedTabs(UCase(CStr(mappedByKey(key)))) = True
          Next

          For r = 1 To rowCount
            key = CStr(colIdx) & ":" & CStr(r)
            If Not mappedByKey.Exists(key) Then
              pick = FirstUnusedCandidate(candidates, usedTabs)

              If Len(pick) = 0 And rowCount = 1 Then
                If ParseTabfieldCode(catTab, catPrefix, catNum) Then
                  If catPrefix = anchorPrefix And (catNum \ 100) = anchorGroup And (catNum Mod 100) <> 0 Then
                    If Not usedTabs.Exists(UCase(catTab)) Then pick = UCase(catTab)
                  End If
                End If
              End If

              If Len(pick) > 0 Then
                mappedByKey(key) = pick
                usedTabs(pick) = True
                inferredItems.Add pick & ":" & CStr(colIdx) & ":" & CStr(r)
                inferredCount = inferredCount + 1
              End If
            End If
          Next
        End If
      End If
    End If
  Next

  Dim finalList
  Set finalList = CreateObject("System.Collections.ArrayList")

  For i = 0 To existingItems.Count - 1
    finalList.Add CStr(existingItems(i))
  Next

  For i = 0 To inferredItems.Count - 1
    finalList.Add CStr(inferredItems(i))
  Next

  BuildEffectiveOutputMap = ArrayListToArray(finalList)
End Function

Function BuildOutputTargets(outItems)
  Dim outTargets: Set outTargets = CreateObject("Scripting.Dictionary")
  Dim i, tgtTab, colIdx, rowIdx

  If ArrayHasElements(outItems) Then
    For i = LBound(outItems) To UBound(outItems)
      If TryParseOutputMapItem(outItems(i), tgtTab, colIdx, rowIdx) Then
        If Not outTargets.Exists(colIdx) Then outTargets.Add colIdx, CreateObject("Scripting.Dictionary")
        If Not outTargets(colIdx).Exists(rowIdx) Then outTargets(colIdx).Add rowIdx, Array()
        outTargets(colIdx)(rowIdx) = PushArray(outTargets(colIdx)(rowIdx), tgtTab)
      End If
    Next
  End If

  Set BuildOutputTargets = outTargets
End Function

Function DetermineRowCount(tsec, haveFilters, filterTabs)
  Dim rowCount: rowCount = 1
  If haveFilters Then rowCount = UBound(filterTabs) - LBound(filterTabs) + 1

  Dim rowLimitKV: rowLimitKV = SafeGet(tsec, "row_limit", "")
  Dim maxRows: maxRows = 0
  If Len(rowLimitKV) > 0 And InStr(rowLimitKV, ",") > 0 Then
    On Error Resume Next
    maxRows = CInt(Split(rowLimitKV, ",")(1))
    On Error GoTo 0
  End If
  If maxRows > 0 And maxRows < rowCount Then rowCount = maxRows

  DetermineRowCount = rowCount
End Function

Function ResolveFilterFragments(rowCount, haveFilters, filterTabs, qAliasNorm, qNorm, learn, LEARN_INI)
  Dim filterFrags()
  ReDim filterFrags(rowCount - 1)

  Dim r, idx, ftabName, ftxt, fFrag, fAcc

  For r = 1 To rowCount
    idx = r - 1
    ftabName = ""
    If haveFilters Then
      ftabName = Trim(CStr(filterTabs(idx)))
    End If

    fFrag = ""
    If Len(ftabName) > 0 Then
      ftxt = CleanAfterColon(TrioCmd("page:get_property " & ftabName))

      If InStr(1, ftxt, "ERROR:", vbTextCompare) > 0 Or LCase(ftabName) = "none" Then
        filterFrags(idx) = ""
      ElseIf Len(Trim(ftxt)) > 0 Then
        If ResolveQualifierChain(ftxt, qAliasNorm, qNorm, learn, fFrag, fAcc) Then
          filterFrags(idx) = fFrag
        Else
          Dim leftText: leftText = Trim(CStr(fAcc))

          If Len(leftText) = 0 Then
            filterFrags(idx) = ""
          Else
            Dim leftKey: leftKey = Replace(leftText, " ", "_")
            Dim sgIsAlias, sgAlias, sgCanon, sgFrag, sgScore, didSg
            didSg = SuggestQualifierMapping(leftKey, qAliasNorm, qNorm, learn, sgIsAlias, sgAlias, sgCanon, sgFrag, sgScore)

            If didSg Then
              If sgIsAlias Then
                AppendIniSectionLine LEARN_INI, "[PENDING_QUALIFIER_ALIASES]", _
                  UCase(leftKey) & " = " & Replace(UCase(sgAlias), "_", " ") & " ; suggest alias QUAL " & ftabName & " score=" & ScoreStr(sgScore)
              Else
                AppendIniSectionLine LEARN_INI, "[PENDING_QUALIFIER]", _
                  UCase(leftKey) & " = " & Replace(UCase(sgCanon), "_", " ") & " ; suggest canon QUAL " & ftabName & " score=" & ScoreStr(sgScore)
              End If

              If Len(Trim(sgFrag)) > 0 Then
                filterFrags(idx) = sgFrag
              Else
                filterFrags(idx) = ""
              End If
            Else
              AppendIniSectionLine LEARN_INI, "[PENDING_QUALIFIER]", UCase(leftKey) & " = (suggest) ??? ; QUAL " & ftabName
              filterFrags(idx) = ""
            End If
          End If
        End If
      Else
        filterFrags(idx) = ""
      End If
    Else
      filterFrags(idx) = ""
    End If
  Next

  ResolveFilterFragments = filterFrags
End Function

Sub ProcessCategoryColumns(catTabs, transforms, rxTransforms, learn, LEARN_INI, catAliasRaw, catRaw, catPAliasRaw, catPRaw, catAliasNorm, catPAliasNorm, catNorm, catPNorm, outTargets, rowCount, filterFrags, qPrefix, qRemFrag, entityCtx, entityType, playerSubtype)
  Dim catConcept, wasPitcher, usedHeur, acceptedBy, fuzzyScore
  Dim col, rawIn, normIn, stripped, preferPitcher, isPitcher, measure, statPath, basePath, fullPath

  For col = LBound(catTabs) To UBound(catTabs)
    rawIn = CleanAfterColon(TrioCmd("page:get_property " & Trim(catTabs(col))))
    normIn = ApplyTransforms(rawIn, transforms, rxTransforms)
    stripped = StripStopwords(normIn, learn("stopwords"))
    preferPitcher = ContainsAnyToken(normIn, learn("prefer_pitcher_tokens"))
    isPitcher = (UCase(entityType) = "PLAYER" And UCase(playerSubtype) = "P")

    If Len(Trim(stripped)) > 0 Then
      measure = ResolveCategoryMeasure(stripped, isPitcher, catAliasNorm, catPAliasNorm, catNorm, catPNorm)

      If Len(measure) = 0 Then
        Dim collapsed : collapsed = CollapseDoubles(stripped)
        If LCase(collapsed) <> LCase(stripped) Then
          measure = ResolveCategoryMeasure(collapsed, isPitcher, catAliasNorm, catPAliasNorm, catNorm, catPNorm)
        End If
      End If

      If Len(measure) = 0 Then
        Dim swapHit : swapHit = TryAdjacentSwapResolve(stripped, isPitcher, catAliasNorm, catPAliasNorm, catNorm, catPNorm)
        If Len(swapHit) > 0 Then
          measure = swapHit
        End If
      End If

      acceptedBy = "" : fuzzyScore = 0
      If Len(measure) = 0 Then
        If ResolveCategorySmart(stripped, preferPitcher, learn, catAliasRaw, catRaw, catPAliasRaw, catPRaw, _
                                statPath, catConcept, wasPitcher, usedHeur, acceptedBy, fuzzyScore) Then
          measure = statPath
        Else
          LogLearnPendingWithGuess LEARN_INI, stripped, catRaw, catPRaw, catAliasRaw, catPAliasRaw, preferPitcher, learn, fuzzyScore
        End If
      End If

      If Len(measure) > 0 Then
        basePath = qPrefix
        If Len(Trim(qRemFrag)) > 0 Then basePath = basePath & "." & qRemFrag

        Dim col1 : col1 = col + 1
        Dim r
        For r = 1 To rowCount
          fullPath = basePath
          Dim idx : idx = r - 1
          If Len(Trim(filterFrags(idx))) > 0 Then fullPath = fullPath & "." & filterFrags(idx)

          Do While InStr(fullPath, "..") > 0
            fullPath = Replace(fullPath, "..", ".")
          Loop

          Dim effMeasure : effMeasure = measure
          Dim useStrippedPath : useStrippedPath = False

          If UCase(effMeasure) = "USAGE" _
             Or UCase(effMeasure) = "_DYNAMIC_ARSENAL_" _
             Or InStr(1, effMeasure, "*", vbTextCompare) > 0 Then

            Dim dyn : dyn = ResolveDynamicUsageMeasure(fullPath)
            If Len(dyn) > 0 Then
              effMeasure = dyn
              useStrippedPath = True
            Else
              effMeasure = "arsenal_percentage"
              useStrippedPath = True
            End If
          End If

          Dim pathForSyntax : pathForSyntax = fullPath
          If useStrippedPath _
            And ( (LCase(Left(effMeasure, 8))  = "arsenal_"        And LCase(Right(effMeasure, 11)) = "_percentage") _
               Or (LCase(Left(effMeasure, 15)) = "pitch_category_" And LCase(Right(effMeasure, 11)) = "_percentage") ) Then
            pathForSyntax = StripPitchFuncs(fullPath)
          End If

          Dim finalSyntax : finalSyntax = "{{stats." & entityCtx & "." & pathForSyntax & "." & effMeasure & "}}"
          finalSyntax = NormalizeSeasonInSyntax(finalSyntax)

          If outTargets.Exists(col1) Then
            If outTargets(col1).Exists(r) Then
              Dim targets, j
              targets = outTargets(col1)(r)
              If IsArray(targets) Then
                For j = LBound(targets) To UBound(targets)
                  Call Tx_SetCustomProp(CStr(targets(j)), finalSyntax)
                Next
              End If
            End If
          End If
        Next
      End If

    Else
      Dim col1_clear : col1_clear = col + 1
      If outTargets.Exists(col1_clear) Then
        'Dim r
        For r = 1 To rowCount
          If outTargets(col1_clear).Exists(r) Then
            Dim tclear, jc
            tclear = outTargets(col1_clear)(r)
            If IsArray(tclear) Then
              For jc = LBound(tclear) To UBound(tclear)
                Call Tx_SetCustomProp(CStr(tclear(jc)), "")
              Next
            End If
          End If
        Next
      End If
    End If
  Next
End Sub

' ---------------- Template execution ----------------
Sub ExecuteTemplatePipeline(srcDir, mappingsIni, LEARN_INI, transforms, rxTransforms, learn)
  On Error Resume Next

  Dim tmplNameRaw: tmplNameRaw = TrioCmd("page:getpagetemplate")
  Dim tmplName: tmplName = NormalizeTemplateNameForKeys(tmplNameRaw)
  If Len(tmplName) = 0 Then
    Call Phase_EarlyExit(PHASE_03_CLASSIFY_FIELDS, "TEMPLATE.EMPTY", "Template name could not be resolved.", "Verify page template binding and reload the page.")
    Call Diag_WriteLine("TX: EARLY EXIT - TEMPLATE_EMPTY")
    Exit Sub
  End If

  Call Phase_Begin(PHASE_03_CLASSIFY_FIELDS, "Classifying template and runtime context")
  Diag_Mark_Classify "Template detected: " & tmplName

  Dim tsec, qualTab, catTabs, outItems, filterTabs, haveFilters
  haveFilters = False
  If Not LoadTemplateSectionConfig(srcDir, tmplName, tsec, qualTab, catTabs, outItems, filterTabs, haveFilters) Then
    Call Phase_Fail(PHASE_03_CLASSIFY_FIELDS, "TPLCFG.MISS", "Template block missing for " & tmplName, "Add [TEMPLATE:" & tmplName & "] to SmartStat_TemplateConfig.ini")
    Exit Sub
  End If
  Diag_Mark_Classify "Template config loaded"

  Dim catRaw, catNorm, catPRaw, catPNorm, catAliasRaw, catAliasNorm, catPAliasRaw, catPAliasNorm
  Dim qRaw, qNorm, qAliasRaw, qAliasNorm
  LoadCategoryAndQualifierMappings mappingsIni, catRaw, catNorm, catPRaw, catPNorm, catAliasRaw, catAliasNorm, catPAliasRaw, catPAliasNorm, qRaw, qNorm, qAliasRaw, qAliasNorm
  Diag_Mark_LoadConfig "Category & qualifier maps loaded"

  Dim entityCtx, entityType, playerSubtype
  DetermineEntityContext entityCtx, entityType, playerSubtype
  Diag_Mark_Classify "Entity context: " & entityCtx & " subtype=" & playerSubtype
  Call Phase_EndOk(PHASE_03_CLASSIFY_FIELDS, "Classification complete")

  Dim qPrefix, qRemFrag
  qPrefix = ""
  qRemFrag = ""
  If Not ProcessQualifier(qualTab, qAliasNorm, qNorm, learn, qPrefix, qRemFrag) Then
    Call Phase_EarlyExit(PHASE_05_DETECT_FILTERS, "QUALIFIER.UNRESOLVED", _
      "Qualifier could not be resolved for tabfield " & CStr(qualTab), _
      "Set the qualifier tabfield to an exact supported token or update learn mappings.")
    Call Diag_OperatorAlert("SmartStat aborted: Qualifier could not be resolved. No changes applied.")
    Call Diag_WriteLine("TX: EARLY EXIT - QUALIFIER_UNRESOLVED_PIPELINE")
    Exit Sub
  End If

  Dim rowCount: rowCount = DetermineRowCount(tsec, haveFilters, filterTabs)
  Call Phase_Begin(PHASE_06_BUILD_OUTMAP, "Building output map")
  Dim inferredOutCount: inferredOutCount = 0
  outItems = BuildEffectiveOutputMap(catTabs, outItems, rowCount, inferredOutCount)

  Dim outTargets: Set outTargets = BuildOutputTargets(outItems)
  If outTargets.Count = 0 Then
    Call Phase_Fail(PHASE_06_BUILD_OUTMAP, "OUTMAP.EMPTY", "output_map resolved to no usable targets for " & tmplName, "Define output_map or ensure category/output tabfields share prefix and hundred-group.")
    Exit Sub
  End If
  Diag_Mark_BuildOutMap "Output targets compiled (inferred=" & CStr(inferredOutCount) & ")"
  Call Phase_EndOk(PHASE_06_BUILD_OUTMAP, "Output map complete")

  Dim filterFrags: filterFrags = ResolveFilterFragments(rowCount, haveFilters, filterTabs, qAliasNorm, qNorm, learn, LEARN_INI)
  Diag_Mark_DetectFilters "RowCount=" & rowCount & ", filters resolved"

  Call Phase_Begin(PHASE_07_BUILD_SYNTAX, "Building syntax for category columns")
  Diag_Mark_BuildSyntax "Building syntax for category columns"
  ProcessCategoryColumns catTabs, transforms, rxTransforms, learn, LEARN_INI, catAliasRaw, catRaw, catPAliasRaw, catPRaw, catAliasNorm, catPAliasNorm, catNorm, catPNorm, outTargets, rowCount, filterFrags, qPrefix, qRemFrag, entityCtx, entityType, playerSubtype
  If TRANSACTION_MODE Then
    Call Diag_WriteLine("TX: Category syntax staged for Trio custom properties (pending commit)")
  Else
    Call Diag_WriteLine("TX: Immediate write mode (TRANSACTION_MODE=False)")
  End If

  Call Phase_EndOk(PHASE_07_BUILD_SYNTAX, "Syntax build complete")
  Diag_Mark_ApplyOverrides "Applying static overrides"
  Call Phase_Begin(PHASE_08_PUSH_TO_TRIO, "Applying static overrides")
  ApplyStaticOverridesByTemplate srcDir & "SmartStat_StaticOverrides.ini", tmplName, UCase(entityCtx), UCase(playerSubtype)
  Call Phase_EndOk(PHASE_08_PUSH_TO_TRIO, "Static override values applied")
  Diag_Mark_PushToTrio "Static override values applied"

  On Error GoTo 0
End Sub

' ---------------- Learn writers ----------------
Sub LogLearnPendingWithGuess(learnPath, key, catMap, catPitchMap, catAlias, catPitchAlias, preferPitcher, learn, fuzzyScore)
  On Error Resume Next
  Dim isPitch, canon, acceptedBy, score
  Dim guess: guess = ""
  Dim did, sec

  did = SuggestCanonKey(key, catMap, catPitchMap, catAlias, catPitchAlias, preferPitcher, learn, canon, isPitch, score, acceptedBy)

  If did Then
    ' VBScript has no IIf — use If/Else
    If CBool(isPitch) Then
      sec = "[PENDING_ALIASES_PITCHER]"
    Else
      sec = "[PENDING_ALIASES]"
    End If

    guess = key & " = " & canon & " ; suggest " & acceptedBy & " score=" & ScoreStr(score)
    AppendIniSectionLine learnPath, sec, guess
  Else
    AppendIniSectionLine learnPath, "[PENDING]", key & " = (suggest) ??? ; last_fuzzy=" & ScoreStr(fuzzyScore)
  End If

  On Error GoTo 0
End Sub

Function SuggestCanonKey(key, catMap, catPitchMap, catAlias, catPitchAlias, preferPitcher, learn, ByRef canonOut, ByRef isPitcher, ByRef scoreOut, ByRef acceptedBy)
  Dim aliasKeys: aliasKeys = MergeKeys(catAlias, catPitchAlias)
  Dim bestA, sA
  If HeuristicPick(key, aliasKeys, learn, bestA, sA) Then
    If catAlias.Exists(bestA) Then canonOut = catAlias(bestA): isPitcher=False: scoreOut=sA: acceptedBy="alias": SuggestCanonKey=True: Exit Function
    If catPitchAlias.Exists(bestA) Then canonOut = catPitchAlias(bestA): isPitcher=True: scoreOut=sA: acceptedBy="alias_pitcher": SuggestCanonKey=True: Exit Function
  End If

  Dim canonKeys: canonKeys = MergeKeys(catMap, catPitchMap)
  Dim bestC, sC
  If HeuristicPick(key, canonKeys, learn, bestC, sC) Then
    If catMap.Exists(bestC) Then canonOut = bestC: isPitcher=False: scoreOut=sC: acceptedBy="canon": SuggestCanonKey=True: Exit Function
    If catPitchMap.Exists(bestC) Then canonOut = bestC: isPitcher=True: scoreOut=sC: acceptedBy="canon_pitcher": SuggestCanonKey=True: Exit Function
  End If

  SuggestCanonKey = False
End Function

' ---------------- Static overrides ----------------
Sub ApplyStaticOverridesByTemplate(staticIniPath, tmplName, entityCtx, playerSubtype)
  On Error Resume Next
  Dim ini
  Dim secSpecific, secGeneric
  Dim sec, sec2
  Dim k, val
  Dim k2, val2
  Dim prevVal

  Set ini = LoadIni(staticIniPath)
  If ini Is Nothing Then
    If DIAG_MODE Then Call Diag_WriteLine("OVERRIDE_APPLY_SKIP ini_unavailable path=" & CStr(staticIniPath))
    Exit Sub
  End If

  secSpecific = "STATIC_FIELD_TO_SYNTAX_" & UCase(entityCtx) & "_" & UCase(tmplName)
  secGeneric  = "STATIC_FIELD_TO_SYNTAX_" & UCase(entityCtx)

  If ini.Exists(secSpecific) Then
    Set sec = ini(secSpecific)
    For Each k In sec.Keys
      val = CStr(sec(k))
      If UCase(entityCtx) = "PLAYER" And UCase(playerSubtype) = "P" Then
        If LCase(val) = "{{info.player.primary_position}}" Then val = "{{info.player.pitcher_hand}}"
      End If
      If DIAG_MODE Then
        prevVal = StaticOverride_GetPrevValue(CStr(k))
        Call Diag_LogOverrideApplied(secSpecific, CStr(k), prevVal, val, entityCtx, playerSubtype)
      End If
      Call Tx_SetCustomProp(CStr(k), val)
    Next
  End If

  If ini.Exists(secGeneric) Then
    Set sec2 = ini(secGeneric)
    For Each k2 In sec2.Keys
      val2 = CStr(sec2(k2))
      If UCase(entityCtx) = "PLAYER" And UCase(playerSubtype) = "P" Then
        If LCase(val2) = "{{info.player.primary_position}}" Then val2 = "{{info.player.pitcher_hand}}"
      End If
      If DIAG_MODE Then
        prevVal = StaticOverride_GetPrevValue(CStr(k2))
        Call Diag_LogOverrideApplied(secGeneric, CStr(k2), prevVal, val2, entityCtx, playerSubtype)
      End If
      Call Tx_SetCustomProp(CStr(k2), val2)
    Next
  End If

  On Error GoTo 0
End Sub

Function StaticOverride_GetPrevValue(ByVal tfName)
  Dim tf
  tf = CStr(tfName)

  If TRANSACTION_MODE Then
    If IsObject(ApplyPlan) Then
      If UCase(TypeName(ApplyPlan)) = "DICTIONARY" Then
        If ApplyPlan.Exists(tf) Then
          StaticOverride_GetPrevValue = CStr(ApplyPlan(tf))
          Exit Function
        End If
      End If
    End If
  End If

  StaticOverride_GetPrevValue = CStr(TrioCmd("tabfield:get_custom_property " & tf))
End Function

Sub Diag_LogOverrideApplied(ByVal sourceSection, ByVal tabfield, ByVal oldValue, ByVal newValue, ByVal entityCtx, ByVal playerSubtype)
  If Not DIAG_MODE Then Exit Sub
  If CStr(oldValue) = CStr(newValue) Then Exit Sub
  Call Diag_WriteLine("OVERRIDE_APPLIED section=" & CStr(sourceSection) & " tabfield=" & CStr(tabfield) & " old=[" & CStr(oldValue) & "] new=[" & CStr(newValue) & "] entityCtx=" & CStr(entityCtx) & " playerSubtype=" & CStr(playerSubtype))
End Sub

' ---------------- Misc ----------------
Function CleanAfterColon(line)
  Dim s, p
  s = Trim(CStr(line))
  p = InStr(s, "=")
  If p > 0 Then s = Mid(s, p + 1)
  s = Replace(Replace(s, vbCr, ""), vbLf, "")
  CleanAfterColon = Trim(s)
End Function

Sub EnsureAmbiguityContext()
  On Error Resume Next

  If (Not IsObject(CompilerContext)) Then
    Set CompilerContext = CreateObject("Scripting.Dictionary")
  ElseIf UCase(TypeName(CompilerContext)) <> "DICTIONARY" Then
    Set CompilerContext = CreateObject("Scripting.Dictionary")
  End If

  If Not CompilerContext.Exists("ambiguous") Then
    Dim amb0: Set amb0 = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    amb0.CompareMode = 1
    On Error GoTo 0
    Set CompilerContext("ambiguous") = amb0
  ElseIf (Not IsObject(CompilerContext("ambiguous"))) Or (UCase(TypeName(CompilerContext("ambiguous"))) <> "DICTIONARY") Then
    Dim amb1: Set amb1 = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    amb1.CompareMode = 1
    On Error GoTo 0
    Set CompilerContext("ambiguous") = amb1
  End If

  On Error GoTo 0
End Sub

Function Ambiguity_SafeTruncate(ByVal s, ByVal maxLen)
  Dim t, m
  t = CStr(s)
  m = CLng(maxLen)
  If m < 0 Then m = 0
  If Len(t) > m Then
    Ambiguity_SafeTruncate = Left(t, m) & "...(len=" & CStr(Len(t)) & ")"
  Else
    Ambiguity_SafeTruncate = t
  End If
End Function

Function Ambiguity_StringifyCandidates(ByVal candidates, ByVal maxItems)
  On Error Resume Next
  Dim cap: cap = CLng(maxItems)
  If cap <= 0 Then cap = 8

  Dim list: Set list = CreateObject("System.Collections.ArrayList")
  Dim i, k, tname

  If IsArray(candidates) Then
    For i = LBound(candidates) To UBound(candidates)
      list.Add Trim(CStr(candidates(i)))
    Next
  ElseIf IsObject(candidates) Then
    tname = UCase(TypeName(candidates))
    If tname = "DICTIONARY" Then
      For Each k In candidates.Keys
        list.Add Trim(CStr(k))
      Next
      list.Sort
    ElseIf InStr(1, tname, "ARRAYLIST", vbTextCompare) > 0 Then
      For i = 0 To candidates.Count - 1
        list.Add Trim(CStr(candidates(i)))
      Next
      list.Sort
    Else
      list.Add Trim(CStr(candidates))
    End If
  Else
    If Len(Trim(CStr(candidates))) > 0 Then list.Add Trim(CStr(candidates))
  End If

  Dim outTxt, take
  outTxt = ""
  take = list.Count
  If take > cap Then take = cap

  For i = 0 To take - 1
    If Len(outTxt) > 0 Then outTxt = outTxt & " | "
    outTxt = outTxt & CStr(list(i))
  Next
  If list.Count > cap Then outTxt = outTxt & " | ... +" & CStr(list.Count - cap)

  Ambiguity_StringifyCandidates = outTxt
  On Error GoTo 0
End Function

Sub Ambiguity_AddDetailed(ByVal decisionLabel, ByVal inputToken, ByVal candidates, ByVal selectedState, ByVal notes, ByVal hint)
  On Error Resume Next
  Call EnsureAmbiguityContext()

  Dim ambHits: Set ambHits = CompilerContext("ambiguous")
  Dim decisionTxt, inputTxt, stateTxt, notesTxt, hintTxt
  Dim entryKey, entryObj, candidateTxt

  decisionTxt = UCase(Trim(CStr(decisionLabel)))
  If Len(decisionTxt) = 0 Then decisionTxt = "UNKNOWN_DECISION"
  inputTxt = Trim(CStr(inputToken))

  stateTxt = UCase(Trim(CStr(selectedState)))
  If Len(stateTxt) = 0 Then stateTxt = "BLOCKED"

  notesTxt = Trim(CStr(notes))

  hintTxt = Trim(CStr(hint))

  entryKey = decisionTxt & "|" & UCase(inputTxt)
  If Not ambHits.Exists(entryKey) Then
    Set entryObj = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    entryObj.CompareMode = 1
    On Error GoTo 0
    Set ambHits(entryKey) = entryObj
  ElseIf IsObject(ambHits(entryKey)) Then
    If UCase(TypeName(ambHits(entryKey))) = "DICTIONARY" Then
      Set entryObj = ambHits(entryKey)
    Else
      Set entryObj = CreateObject("Scripting.Dictionary")
      Set ambHits(entryKey) = entryObj
    End If
  Else
    Set entryObj = CreateObject("Scripting.Dictionary")
    Set ambHits(entryKey) = entryObj
  End If

  entryObj("decision_key") = decisionTxt
  entryObj("input_token") = inputTxt
  candidateTxt = Ambiguity_StringifyCandidates(candidates, 8)
  entryObj("candidates") = candidateTxt
  entryObj("state") = stateTxt
  entryObj("notes") = notesTxt
  entryObj("hint") = hintTxt
  If stateTxt = "BLOCKED" Or stateTxt = "FAILED" Or stateTxt = "UNRESOLVED" Then
    entryObj("blocked") = "True"
  Else
    entryObj("blocked") = "False"
  End If

  On Error GoTo 0
End Sub

Sub Diag_WriteAmbiguitySummary()
  On Error Resume Next
  If Not CBool(DIAG_MODE) Then Exit Sub
  If CBool(G_AMBIGUITY_SUMMARY_EMITTED) Then Exit Sub

  Call EnsureAmbiguityContext()
  If (CompilerContext Is Nothing) Then Exit Sub
  Dim allowAmbSummary: allowAmbSummary = False
  If CompilerContext.Exists("learn") Then
    Dim learnTypeNameSummary: learnTypeNameSummary = TypeName(CompilerContext("learn"))
    If IsObject(CompilerContext("learn")) And UCase(CStr(learnTypeNameSummary)) = "DICTIONARY" Then
      Dim lkSummary: Set lkSummary = CompilerContext("learn")
      If lkSummary.Exists("allow_ambiguous_apply") Then allowAmbSummary = CBool(lkSummary("allow_ambiguous_apply"))
    End If
  End If
  If Not allowAmbSummary Then Exit Sub
  If Not CompilerContext.Exists("ambiguous") Then Exit Sub

  Dim ambHits: Set ambHits = CompilerContext("ambiguous")
  If (ambHits Is Nothing) Then Exit Sub
  If ambHits.Count <= 0 Then Exit Sub

  Dim keyList: Set keyList = CreateObject("System.Collections.ArrayList")
  Dim k
  For Each k In ambHits.Keys
    keyList.Add CStr(k)
  Next
  keyList.Sort

  Call Diag_WriteLine("AMBIGUITY_SUMMARY")

  Dim i, hitKey, hitVal, dKey, inTok, cand, st, rsn, act, lineTxt
  For i = 0 To keyList.Count - 1
    hitKey = CStr(keyList(i))
    dKey = "UNKNOWN_DECISION"
    inTok = ""
    cand = ""
    st = "BLOCKED"
    rsn = ""
    act = ""

    If IsObject(ambHits(hitKey)) Then
      If UCase(TypeName(ambHits(hitKey))) = "DICTIONARY" Then
        Set hitVal = ambHits(hitKey)
        If hitVal.Exists("decision_key") Then dKey = CStr(hitVal("decision_key"))
        If hitVal.Exists("input_token") Then inTok = CStr(hitVal("input_token"))
        If hitVal.Exists("candidates") Then cand = CStr(hitVal("candidates"))
        If hitVal.Exists("state") Then st = CStr(hitVal("state"))
        If hitVal.Exists("notes") Then rsn = CStr(hitVal("notes"))
        If hitVal.Exists("hint") Then act = CStr(hitVal("hint"))
      Else
        rsn = CStr(ambHits(hitKey))
      End If
    Else
      rsn = CStr(ambHits(hitKey))
    End If

    If Len(cand) = 0 Then cand = "(none recorded)"
    If Len(act) = 0 Then
      act = "Use an exact token in the template tabfield or add/adjust learn alias to force one match."
    End If

    lineTxt = "AMBIGUITY[" & CStr(i + 1) & "] decision=" & dKey & _
              " input=[" & Ambiguity_SafeTruncate(inTok, 80) & "]" & _
              " state=" & st & _
              " candidates=[" & Ambiguity_SafeTruncate(cand, 220) & "]" & _
              " reason=[" & Ambiguity_SafeTruncate(rsn, 220) & "]" & _
              " action=[" & Ambiguity_SafeTruncate(act, 220) & "]"
    Call Diag_WriteLine(lineTxt)
  Next

  G_AMBIGUITY_SUMMARY_EMITTED = True
  On Error GoTo 0
End Sub

Function BuildAmbiguityOperatorSection()
  On Error Resume Next
  Call EnsureAmbiguityContext

  BuildAmbiguityOperatorSection = ""

  If (CompilerContext Is Nothing) Then Exit Function
  Dim allowAmbOut: allowAmbOut = False
  If CompilerContext.Exists("learn") Then
    Dim learnTypeNameOut: learnTypeNameOut = TypeName(CompilerContext("learn"))
    If IsObject(CompilerContext("learn")) And UCase(CStr(learnTypeNameOut)) = "DICTIONARY" Then
      Dim lkOut: Set lkOut = CompilerContext("learn")
      If lkOut.Exists("allow_ambiguous_apply") Then allowAmbOut = CBool(lkOut("allow_ambiguous_apply"))
    End If
  End If
  If Not allowAmbOut Then Exit Function
  If Not CompilerContext.Exists("ambiguous") Then Exit Function

  Dim amb: Set amb = CompilerContext("ambiguous")
  If (amb Is Nothing) Then Exit Function
  If amb.Count <= 0 Then Exit Function

  Dim keys: keys = amb.Keys
  Dim i, j, tmp
  For i = 0 To UBound(keys) - 1
    For j = i + 1 To UBound(keys)
      If StrComp(CStr(keys(i)), CStr(keys(j)), vbTextCompare) > 0 Then
        tmp = keys(i)
        keys(i) = keys(j)
        keys(j) = tmp
      End If
    Next
  Next

  Dim out: out = "AMBIGUITY:"
  For i = 0 To UBound(keys)
    out = out & vbCrLf & CStr(amb(CStr(keys(i))))
  Next

  BuildAmbiguityOperatorSection = out

  On Error GoTo 0
End Function

' ==========================================
' v4.0 Phase 3: Harness + Integrity Engine
' ==========================================
Function Harness_GetModeFromControlA()
  On Error Resume Next
  Harness_GetModeFromControlA = "OFF"
  If Not HARNESS_ENABLE Then Exit Function

  ' IMPORTANT: match existing control behavior:
  ' - Prefer operator-visible A value (page:get_property A)
  ' - Fallback to custom property A
  Dim v: v = UCase(Trim(CStr(TrioCmd("page:get_property " & HARNESS_CONTROL_TAB))))
  If Len(v) = 0 Then v = UCase(Trim(CStr(TrioCmd("tabfield:get_custom_property " & HARNESS_CONTROL_TAB))))
  If Len(v) = 0 Then Exit Function

  ' Must be before HARNESS_CAPTURE because of substring overlap
  If v = HARNESS_MODE_CAPTURE_ONLY Then Harness_GetModeFromControlA = HARNESS_MODE_CAPTURE_ONLY : Exit Function
  If InStr(v, "SMARTSTAT=" & HARNESS_MODE_CAPTURE_ONLY) > 0 Then Harness_GetModeFromControlA = HARNESS_MODE_CAPTURE_ONLY : Exit Function

  If InStr(v, "SMARTSTAT=HARNESS_CAPTURE") > 0 Then Harness_GetModeFromControlA = "HARNESS_CAPTURE" : Exit Function
  If InStr(v, "SMARTSTAT=HARNESS_STRICT") > 0 Then Harness_GetModeFromControlA = "HARNESS_STRICT" : Exit Function
  If InStr(v, "SMARTSTAT=HARNESS_COMMIT") > 0 Then Harness_GetModeFromControlA = "HARNESS_COMMIT" : Exit Function
  If InStr(v, "SMARTSTAT=HARNESS") > 0 Then Harness_GetModeFromControlA = "HARNESS" : Exit Function
End Function

Sub Harness_EnsureHarnessDir()
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(HARNESS_DIR) Then fso.CreateFolder HARNESS_DIR
  On Error GoTo 0
End Sub

Sub Harness_SnapshotPageState(ByRef outVisible, ByRef outCustomProps)
  On Error Resume Next
  Dim tabs, arr, t, v, cp
  tabs = TrioCmd("page:get_tabfield_names")
  arr = Split(CStr(tabs))

  For Each t In arr
    v = CStr(TrioCmd("page:get_property " & CStr(t)))
    cp = CStr(TrioCmd("tabfield:get_custom_property " & CStr(t)))
    outVisible(CStr(t)) = v
    outCustomProps(CStr(t)) = cp
  Next
  On Error GoTo 0
End Sub

Sub Harness_ResetPostState()
  Set G_HARNESS_POST_CP = Nothing
  Set G_HARNESS_POST_V = Nothing
  G_HARNESS_POST_SKIPPED_REASON = ""
End Sub

Sub Harness_MarkPostSkipped(ByVal reason)
  Set G_HARNESS_POST_CP = Nothing
  Set G_HARNESS_POST_V = Nothing
  G_HARNESS_POST_SKIPPED_REASON = UCase(Trim(CStr(reason)))
End Sub

Sub Harness_FinalizeEarlyExitArtifacts(ByVal reason)
  On Error Resume Next
  If Not HARNESS_ENABLE Then Exit Sub
  If UCase(Trim(CStr(G_HARNESS_MODE))) = "OFF" Then Exit Sub
  If CBool(G_HARNESS_EARLY_EXIT_FINALIZED) Then Exit Sub

  G_HARNESS_EARLY_EXIT_FINALIZED = True
  Call Harness_CapturePostSnapshot()
  Call Harness_WriteSnapshotArtifact(CStr(reason))
  Call Harness_WriteGroupedDiffArtifact(CStr(reason))
  On Error GoTo 0
End Sub

Sub Harness_CapturePostSnapshot()
  On Error Resume Next
  Dim cp, vv
  Set cp = NewTextDict()
  Set vv = NewTextDict()
  Call Harness_SnapshotPageState(vv, cp)
  If Err.Number = 0 Then
    Set G_HARNESS_POST_CP = cp
    Set G_HARNESS_POST_V = vv
    G_HARNESS_POST_SKIPPED_REASON = ""
  Else
    Call Harness_MarkPostSkipped("SNAPSHOT_ERROR")
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Function Harness_IsDict(ByVal d)
  Harness_IsDict = False
  If IsObject(d) Then
    If UCase(TypeName(d)) = "DICTIONARY" Then Harness_IsDict = True
  End If
End Function

Function Harness_SortedKeys(ByVal d)
  On Error Resume Next
  Dim list, k
  Set list = CreateObject("System.Collections.ArrayList")

  If Err.Number = 0 And Not (list Is Nothing) Then
    If Harness_IsDict(d) Then
      For Each k In d.Keys
        list.Add CStr(k)
      Next
    End If
    list.Sort
    Set Harness_SortedKeys = list
    On Error GoTo 0
    Exit Function
  End If

  ' Fallback: preserve unsorted key order in a dictionary-backed index list.
  Err.Clear
  Dim fallback, idx
  Set fallback = NewTextDict()
  idx = 0
  If Harness_IsDict(d) Then
    For Each k In d.Keys
      fallback(CStr(idx)) = CStr(k)
      idx = idx + 1
    Next
  End If
  Set Harness_SortedKeys = fallback
  On Error GoTo 0
End Function

Function Harness_DictHasKey(ByVal d, ByVal k)
  Harness_DictHasKey = False
  If Harness_IsDict(d) Then
    Harness_DictHasKey = d.Exists(CStr(k))
  End If
End Function

Function Harness_DictGetSafe(ByVal d, ByVal k)
  Harness_DictGetSafe = ""
  If Harness_DictHasKey(d, k) Then
    Harness_DictGetSafe = CStr(d(CStr(k)))
  End If
End Function

Function Harness_UnionSortedKeys(ByVal d1, ByVal d2)
  Dim merged, k
  Set merged = NewTextDict()
  If Harness_IsDict(d1) Then
    For Each k In d1.Keys
      merged(CStr(k)) = True
    Next
  End If
  If Harness_IsDict(d2) Then
    For Each k In d2.Keys
      merged(CStr(k)) = True
    Next
  End If
  Set Harness_UnionSortedKeys = Harness_SortedKeys(merged)
End Function

Function Harness_KeyListCount(ByVal keysObj)
  Harness_KeyListCount = 0
  If Not IsObject(keysObj) Then Exit Function
  Select Case UCase(TypeName(keysObj))
    Case "ARRAYLIST", "DICTIONARY"
      Harness_KeyListCount = CLng(keysObj.Count)
  End Select
End Function

Function Harness_KeyListItem(ByVal keysObj, ByVal idx)
  Harness_KeyListItem = ""
  If Not IsObject(keysObj) Then Exit Function
  Select Case UCase(TypeName(keysObj))
    Case "ARRAYLIST"
      Harness_KeyListItem = CStr(keysObj(CLng(idx)))
    Case "DICTIONARY"
      If keysObj.Exists(CStr(idx)) Then Harness_KeyListItem = CStr(keysObj(CStr(idx)))
  End Select
End Function

Sub Harness_WriteSnapshotSection(ByRef ts, ByVal sectionName, ByVal dataDict)
  Dim keysObj, i, n, k, lineVal
  ts.WriteLine "[" & CStr(sectionName) & "]"
  Set keysObj = Harness_SortedKeys(dataDict)
  n = Harness_KeyListCount(keysObj)
  If n = 0 Then
    ts.WriteLine "(none)"
  Else
    For i = 0 To (n - 1)
      k = Harness_KeyListItem(keysObj, i)
      lineVal = CStr(Harness_DictGetSafe(dataDict, k))
      lineVal = Replace(Replace(CStr(lineVal), vbCrLf, "\n"), vbTab, " ")
      ts.WriteLine CStr(k) & "=" & CStr(lineVal)
    Next
  End If
  ts.WriteLine ""
End Sub

Sub Harness_WriteSnapshotArtifact(ByVal reason)
  On Error Resume Next
  Dim runId, p, fso, ts, postAvailable, postReason

  If Not HARNESS_ENABLE Then Exit Sub
  If UCase(Trim(CStr(G_HARNESS_MODE))) = "OFF" Then Exit Sub

  runId = Trim(CStr(gDiagRunId))
  If Len(runId) = 0 Then runId = Diag_TimestampCompact()
  p = HARNESS_DIR & "Harness_Snapshot_" & CStr(runId) & ".txt"

  Call Harness_EnsureHarnessDir()
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.CreateTextFile(p, True)
  If ts Is Nothing Then Exit Sub

  ts.WriteLine "RUN_ID=" & CStr(runId)
  ts.WriteLine "TEMPLATE=" & CStr(gTemplateName)
  ts.WriteLine "MODE=" & CStr(G_HARNESS_MODE)
  ts.WriteLine "REASON=" & CStr(reason)
  ts.WriteLine ""

  Call Harness_WriteSnapshotSection(ts, "PRE.CP", G_HARNESS_PRE_CP)
  Call Harness_WriteSnapshotSection(ts, "PRE.VALUES", G_HARNESS_PRE_V)

  postAvailable = (Harness_IsDict(G_HARNESS_POST_CP) And Harness_IsDict(G_HARNESS_POST_V))
  If postAvailable Then
    Call Harness_WriteSnapshotSection(ts, "POST.CP", G_HARNESS_POST_CP)
    Call Harness_WriteSnapshotSection(ts, "POST.VALUES", G_HARNESS_POST_V)
  Else
    postReason = Trim(CStr(G_HARNESS_POST_SKIPPED_REASON))
    If Len(postReason) = 0 Then postReason = "UNKNOWN"
    ts.WriteLine "POST_SKIPPED reason=" & CStr(postReason)
    ts.WriteLine ""
  End If

  ts.Close
  Set ts = Nothing
  If DIAG_MODE Then Call Diag_WriteLine("HARNESS: snapshot artifact written: " & p)
  On Error GoTo 0
End Sub

Sub Harness_WriteGroupedDiffArtifact(ByVal reason)
  On Error Resume Next
  Dim runId, p, fso, ts, postAvailable, postReason
  Dim keysObj, i, n, k
  Dim hasBefore, hasAfter, beforeV, afterV
  Dim cpChanges, valueChanges

  If Not HARNESS_ENABLE Then Exit Sub
  If UCase(Trim(CStr(G_HARNESS_MODE))) = "OFF" Then Exit Sub

  runId = Trim(CStr(gDiagRunId))
  If Len(runId) = 0 Then runId = Diag_TimestampCompact()
  p = HARNESS_DIR & "Harness_Diff_" & CStr(runId) & ".txt"

  Call Harness_EnsureHarnessDir()
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.CreateTextFile(p, True)
  If ts Is Nothing Then Exit Sub

  ts.WriteLine "RUN_ID=" & CStr(runId)
  ts.WriteLine "TEMPLATE=" & CStr(gTemplateName)
  ts.WriteLine "MODE=" & CStr(G_HARNESS_MODE)
  ts.WriteLine "REASON=" & CStr(reason)
  ts.WriteLine ""

  postAvailable = (Harness_IsDict(G_HARNESS_POST_CP) And Harness_IsDict(G_HARNESS_POST_V))
  If Not postAvailable Then
    postReason = Trim(CStr(G_HARNESS_POST_SKIPPED_REASON))
    If Len(postReason) = 0 Then postReason = "UNKNOWN"
    ts.WriteLine "POST_SKIPPED reason=" & CStr(postReason)
    ts.WriteLine ""
  End If

  cpChanges = 0
  valueChanges = 0

  ts.WriteLine "[CP_CHANGES]"
  If postAvailable Then
    Set keysObj = Harness_UnionSortedKeys(G_HARNESS_PRE_CP, G_HARNESS_POST_CP)
    n = Harness_KeyListCount(keysObj)
    For i = 0 To (n - 1)
      k = Harness_KeyListItem(keysObj, i)
      hasBefore = Harness_DictHasKey(G_HARNESS_PRE_CP, k)
      hasAfter = Harness_DictHasKey(G_HARNESS_POST_CP, k)
      If hasBefore Then beforeV = CStr(Harness_DictGetSafe(G_HARNESS_PRE_CP, k)) Else beforeV = "<MISSING>"
      If hasAfter Then afterV = CStr(Harness_DictGetSafe(G_HARNESS_POST_CP, k)) Else afterV = "<MISSING>"
      If CStr(beforeV) <> CStr(afterV) Then
        cpChanges = cpChanges + 1
        beforeV = Replace(Replace(CStr(beforeV), vbCrLf, "\n"), vbTab, " ")
        afterV = Replace(Replace(CStr(afterV), vbCrLf, "\n"), vbTab, " ")
        ts.WriteLine CStr(k) & ": [" & CStr(beforeV) & "] -> [" & CStr(afterV) & "]"
      End If
    Next
  End If
  If cpChanges = 0 Then ts.WriteLine "(none)"
  ts.WriteLine ""

  ts.WriteLine "[VALUE_CHANGES]"
  If postAvailable Then
    Set keysObj = Harness_UnionSortedKeys(G_HARNESS_PRE_V, G_HARNESS_POST_V)
    n = Harness_KeyListCount(keysObj)
    For i = 0 To (n - 1)
      k = Harness_KeyListItem(keysObj, i)
      hasBefore = Harness_DictHasKey(G_HARNESS_PRE_V, k)
      hasAfter = Harness_DictHasKey(G_HARNESS_POST_V, k)
      If hasBefore Then beforeV = CStr(Harness_DictGetSafe(G_HARNESS_PRE_V, k)) Else beforeV = "<MISSING>"
      If hasAfter Then afterV = CStr(Harness_DictGetSafe(G_HARNESS_POST_V, k)) Else afterV = "<MISSING>"
      If CStr(beforeV) <> CStr(afterV) Then
        valueChanges = valueChanges + 1
        beforeV = Replace(Replace(CStr(beforeV), vbCrLf, "\n"), vbTab, " ")
        afterV = Replace(Replace(CStr(afterV), vbCrLf, "\n"), vbTab, " ")
        ts.WriteLine CStr(k) & ": [" & CStr(beforeV) & "] -> [" & CStr(afterV) & "]"
      End If
    Next
  End If
  If valueChanges = 0 Then ts.WriteLine "(none)"
  ts.WriteLine ""

  G_HARNESS_DIFF_COUNT = CLng(cpChanges) + CLng(valueChanges)
  ts.WriteLine "DIFF_COUNT=" & CStr(G_HARNESS_DIFF_COUNT)

  ts.Close
  Set ts = Nothing
  If DIAG_MODE Then Call Diag_WriteLine("HARNESS: grouped diff written: " & p & " (diffCount=" & CStr(G_HARNESS_DIFF_COUNT) & ")")
  On Error GoTo 0
End Sub

Sub Harness_WriteFixtureFile(ByVal visDict, ByVal cpDict, ByVal tmplName)
  On Error Resume Next
  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim p: p = HARNESS_DIR & "fixture_" & NormalizeKey(CStr(tmplName)) & "_" & CStr(gDiagRunId) & ".ini"
  Dim ts: Set ts = fso.CreateTextFile(p, True)

  ts.WriteLine "[META]"
  ts.WriteLine "template=" & CStr(tmplName)
  ts.WriteLine "run_id=" & CStr(gDiagRunId)
  ts.WriteLine "machine=" & Diag_SafeEnv("COMPUTERNAME")
  ts.WriteLine "user=" & Diag_SafeEnv("USERNAME")
  ts.WriteLine ""

  ts.WriteLine "[VISIBLE]"
  Dim k
  For Each k In visDict.Keys
    ts.WriteLine CStr(k) & "=" & Replace(CStr(visDict(k)), vbCrLf, "\n")
  Next
  ts.WriteLine ""

  ts.WriteLine "[CUSTOM_PROPERTIES]"
  For Each k In cpDict.Keys
    ts.WriteLine CStr(k) & "=" & Replace(CStr(cpDict(k)), vbCrLf, "\n")
  Next

  ts.Close
  Set ts = Nothing
  Call Diag_WriteLine("HARNESS: fixture written: " & p)
  On Error GoTo 0
End Sub

Sub Harness_WriteIntegrityDiff(ByVal preCustomProps, ByVal planDict)
  On Error Resume Next

  ' Reset diff count every run (used by HARNESS_STRICT commit gate)
  G_HARNESS_DIFF_COUNT = 0

  Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  Dim p: p = HARNESS_DIR & "diff_" & NormalizeKey(CStr(gTemplateName)) & "_" & CStr(gDiagRunId) & ".txt"
  Dim ts: Set ts = fso.CreateTextFile(p, True)

  ts.WriteLine "SMARTSTAT HARNESS DIFF"
  ts.WriteLine "template=" & CStr(gTemplateName)
  ts.WriteLine "run_id=" & CStr(gDiagRunId)
  ts.WriteLine "mode=" & CStr(G_HARNESS_MODE)
  ts.WriteLine "planned_fields=" & IIf(planDict Is Nothing, "0", CStr(planDict.Count))
  ts.WriteLine String(60, "-")

  If (planDict Is Nothing) Or (planDict.Count = 0) Then
    ts.WriteLine "(no planned changes)"
    ts.WriteLine ""
    ts.WriteLine "DIFF_COUNT=0"
    ts.Close
    Call Diag_WriteLine("HARNESS: integrity diff written: " & p & " (diffCount=0)")
    Exit Sub
  End If

  Dim k, beforeV, afterV
  For Each k In planDict.Keys
    beforeV = ""
    If Not (preCustomProps Is Nothing) Then
      If preCustomProps.Exists(CStr(k)) Then beforeV = CStr(preCustomProps(CStr(k)))
    End If
    afterV = CStr(planDict(CStr(k)))

    If CStr(beforeV) <> CStr(afterV) Then
      G_HARNESS_DIFF_COUNT = G_HARNESS_DIFF_COUNT + 1
      ts.WriteLine "[" & CStr(k) & "]"
      ts.WriteLine "  BEFORE: " & Replace(Replace(CStr(beforeV), vbCrLf, "\n"), vbTab, " ")
      ts.WriteLine "  AFTER : " & Replace(Replace(CStr(afterV), vbCrLf, "\n"), vbTab, " ")
    End If
  Next

  ts.WriteLine ""
  ts.WriteLine "DIFF_COUNT=" & CStr(G_HARNESS_DIFF_COUNT)

  ts.Close
  Set ts = Nothing
  Call Diag_WriteLine("HARNESS: integrity diff written: " & p & " (diffCount=" & CStr(G_HARNESS_DIFF_COUNT) & ")")
  On Error GoTo 0
End Sub

' ---------------- Socket refresh ----------------
Function SmartStat_RefreshSocketData()
  Dim tabs, tab_arr, tab, flag
  Dim on_air_tabs, oa_tab
  Dim page_name, page_desc

  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  on_air_tabs = "["

  For Each tab In tab_arr
    flag = TrioCmd("tabfield:get_custom_property " & tab)

    If flag <> "" Then
      ' Skip tabs with trivial custom props
      If Left(flag, 4) <> "SMT=" And _
         Left(flag, 1) <> "x" And _
         UCase(Left(flag, 5)) <> "HOME " And _
         UCase(Left(flag, 5)) <> "AWAY " And _
         UCase(Left(flag, 7)) <> "TEAM XX" And _
         UCase(Left(flag, 9)) <> "PLAYER XX" Then

         oa_tab = "['" & tab & "','" & flag & "'], "
         on_air_tabs = on_air_tabs & oa_tab
      End If
    End If
  Next

  on_air_tabs = on_air_tabs & "]"

  If on_air_tabs <> "[]" Then
    page_name = TrioCmd("page:getpagename")
    page_desc = TrioCmd("page:getpagedescription")

    Dim ambOut: ambOut = BuildAmbiguityOperatorSection()
    If Len(ambOut) > 0 Then
      If Len(page_desc) > 0 Then
        page_desc = page_desc & vbCrLf & ambOut
      Else
        page_desc = ambOut
      End If
    End If

    page_desc = Replace(page_desc, vbCrLf, "\n")
    page_desc = Replace(page_desc, vbCr, "\n")
    page_desc = Replace(page_desc, vbLf, "\n")

    If TrioCmd("sock:socket_is_connected") Then
      TrioCmd "sock:send_socket_data on_air_get message_number=" & page_name & _
              " query=" & on_air_tabs & " message_context=" & page_desc & vbCrLf
    End If
  End If
End Function
