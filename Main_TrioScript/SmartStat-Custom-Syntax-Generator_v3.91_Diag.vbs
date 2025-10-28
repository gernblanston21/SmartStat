' SmartStat-Custom-Syntax-Generator_v3.91_Learn.vbs
Option Explicit
' =========================
' SmartStat Operator Diagnostics (v1.0)
' =========================
' Purpose: Pinpoint exactly where SmartStat fails on some operator machines.

Const DIAG_MODE                = False
Const DIAG_LOG_DIR             = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\"
Const DIAG_PREFIX              = "SmartStat_OperatorDiag_"
Const DIAG_ENV_PREFIX          = "SmartStat_EnvCheck_"
Const SMARTSTAT_DEBUG_FILE     = "SmartStat_Debug.txt"
Const DIAG_MAX_FILE_SIZE_BYTES = "5242880"

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
    Diag_Step PHASE_01_ENV_VALIDATE, "Validating environment and prerequisites"
    Dim ok
    ok = Diag_Assert(Diag_CheckWSHEnabled(), PHASE_01_ENV_VALIDATE, "ENV.WSH", "Windows Script Host appears enabled", "Enable: HKLM\Software\Microsoft\Windows Script Host\Settings\Enabled=1 (or key absent)")
    If Not ok Then Diag_Check_Environment = False : Exit Function
    ok = Diag_Assert(Diag_CheckWriteAccess(DIAG_LOG_DIR), PHASE_01_ENV_VALIDATE, "ENV.LOGWRITE", "Write access to " & DIAG_LOG_DIR, "Grant write perms to " & DIAG_LOG_DIR & " or choose a writable folder")
    If Not ok Then Diag_Check_Environment = False : Exit Function
    Dim adoOK : adoOK = Diag_CheckADOAvailable()
    If adoOK Then
        Diag_WriteLine "[OK] ENV.ADO: ADODB present"
    Else
        Diag_WriteLine "[WARN] ENV.ADO: ADODB registry not found; if SmartStat uses ADO on this page, install MDAC/ADO."
    End If
    Diag_Check_Environment = True
End Function

Function Diag_Check_ConfigPresence(ByVal mappingsPath, ByVal overridesPath, ByVal templateCfgPath)
    If Not DIAG_MODE Then Diag_Check_ConfigPresence = True : Exit Function
    Diag_Step PHASE_02_LOAD_CONFIG, "Probing config files"
    Dim ok
    ok = Diag_Assert(Diag_FileExistsReadable(mappingsPath), PHASE_02_LOAD_CONFIG, "CFG.MAP", "Mappings file readable: " & mappingsPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then Diag_Check_ConfigPresence = False : Exit Function
    ok = Diag_Assert(Diag_FileExistsReadable(overridesPath), PHASE_02_LOAD_CONFIG, "CFG.OVR", "Static overrides file readable: " & overridesPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then Diag_Check_ConfigPresence = False : Exit Function
    ok = Diag_Assert(Diag_FileExistsReadable(templateCfgPath), PHASE_02_LOAD_CONFIG, "CFG.TPL", "Template config file readable: " & templateCfgPath, "Verify path and permissions. Redownload if corrupted.")
    If Not ok Then Diag_Check_ConfigPresence = False : Exit Function
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

  
  Dim x_tmplForDiag: x_tmplForDiag = TrioCmd("page:getpagetemplate")
  Call Diag_Init(x_tmplForDiag)
  If Not Diag_Check_Environment() Then Exit Sub
  Dim LOG_FILE:      LOG_FILE      = "E:\EDRIVE\UNIVERSAL\SmartStat\DiagLogs\SmartStat_LearnDebug.txt"
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
    Call Diag_Done()
    Exit Sub
  End If

  Diag_Mark_LoadConfig "Loading INI mappings"
  EnsureLogDir LOG_FILE
  EnsureLogDir DIAG_LOG_DIR & "._probe"
  LogLine LOG_FILE, "=== v3.91_Learn START ==="
  LogLine LOG_FILE, "INFO: Using mappings at [" & MAPPINGS_INI & "]"
  LogLine LOG_FILE, "INFO: Learn file path  [" & LEARN_INI & "]"

  Dim ini: Set ini = LoadIni(MAPPINGS_INI)
  If ini Is Nothing Then
    GuiErr "Failed to load INI: " & MAPPINGS_INI
    LogLine LOG_FILE, "ERROR: INI load failed (" & MAPPINGS_INI & ")"

    ' still run so static overrides etc. can apply
    ExecuteTemplatePipeline SRC_DIR, Nothing, LEARN_INI, NewTextDict(), NewTextDict(), NewTextDict()

    ' finalize and return (no GoTo)
    Call Diag_Done()
    Call FinalizeAndRefresh(LOG_FILE, startTime)
    Exit Sub
  End If

  Dim transforms:   Set transforms   = LoadTransforms(ini)
  Dim rxTransforms: Set rxTransforms = LoadRegexDict(ini, "TRANSFORMS_REGEX")
  Dim learn:        Set learn        = LoadLearnKnobs(ini)

  ExecuteTemplatePipeline SRC_DIR, ini, LEARN_INI, transforms, rxTransforms, learn

  ' normal finalize (no label)
  Call Diag_Done()
  Call FinalizeAndRefresh(LOG_FILE, startTime)
End Sub

Sub FinalizeAndRefresh(logFile, startT)
  Dim elapsed: elapsed = Round(Timer - startT, 2)
  LogLine logFile, "SCRIPT RUNTIME: " & elapsed & " seconds"
  LogLine logFile, "=== v3.91_Learn END ==="
  LogLine logFile, " "
  On Error GoTo 0
  Call SmartStat_RefreshSocketData()
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
' Strip UTF-8 BOM if present on first line
If Not tf.AtEndOfStream Then
    Dim firstLinePos: firstLinePos = tf.Line
    Dim peekLine: peekLine = tf.ReadLine
    If Len(peekLine) > 0 Then
        If AscW(Left(peekLine,1)) = &HFEFF Then peekLine = Mid(peekLine,2)
    End If
    tf.Close
    Set tf = fso.OpenTextFile(path, 1, False)
End If

  Dim dict: Set dict = NewTextDict()
  Dim sec, line, eq, k, v
  sec = ""
  Do Until tf.AtEndOfStream
    line = Trim(tf.ReadLine)
    If Len(line) = 0 Then
    ElseIf Left(line,1) = ";" Or Left(line,1) = "#" Then
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
  src = "áàäâãåčçďéèëêěíìïîľĺńñóòöôõřśšťúùüûýžÁÀÄÂÃÅČÇĎÉÈËÊĚÍÌÏÎĽĹŃÑÓÒÖÔÕŘŚŠŤÚÙÜÛÝŽ"
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
  d("stopwords") = Array("batting","bat","hitting","offensive","rate","percent","percentage","team")
  d("deny_fuzzy") = Array("OBP","OPS","ERA","WHIP","WAR")
  d("prefer_pitcher_tokens") = Array("fb","velo","spin","whiff","csw")
  If ini.Exists("LEARN") Then
    Dim sec: Set sec = ini("LEARN")
    If sec.Exists("fuzzy_threshold") Then d("fuzzy_threshold") = CDbl(sec("fuzzy_threshold"))
    If sec.Exists("fuzzy_threshold_short") Then d("fuzzy_threshold_short") = CDbl(sec("fuzzy_threshold_short"))
    If sec.Exists("phonetic_enable") Then d("phonetic_enable") = (LCase(sec("phonetic_enable"))="true")
    If sec.Exists("jaccard_threshold") Then d("jaccard_threshold") = CDbl(sec("jaccard_threshold"))
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

  Dim bestA, sA
  If HeuristicPick(keyN, qAliasNorm.Keys, learn, bestA, sA) Then
    If qAliasNorm.Exists(bestA) Then
      Dim canon: canon = NormalizeKey(CStr(qAliasNorm(bestA)))
      If qNorm.Exists(canon) Then outFrag = CStr(qNorm(canon)) : acceptedBy="fuzzy/alias" : scoreOut=sA : ResolveQualifierSmart=True : Exit Function
    End If
  End If

  Dim bestC, sC
  If HeuristicPick(keyN, qNorm.Keys, learn, bestC, sC) Then
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
  Dim bestAlias, s1
  If HeuristicPick(keyTrim, aliasKeys, learn, bestAlias, s1) Then
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
  Dim bestCanon, s2
  If HeuristicPick(keyTrim, canonKeys, learn, bestCanon, s2) Then
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

Private Sub ScanForBest(qn, arr, ByRef bestKey, ByRef bestDist)
  Dim i, k, dist
  For i = LBound(arr) To UBound(arr)
    k = CStr(arr(i))
    dist = Lev(qn, k)
    If dist < bestDist Then
      bestDist = dist
      bestKey = k
      If Len(qn) <= 5 And bestDist <= 1 Then Exit Sub
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
  tf.WriteLine Now() & " [v3.91_Learn] " & s
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

  fragJoined = ""                              ' e.g., "rspos(true).vs_pitcher_hand(left).inning(7-9)"
  leftoversText = ""                           ' any unmatched tokens (for learn logging)

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
		  ' try alias -> canon with singular
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

  ResolveQualifierChain = (Len(fragJoined) > 0)
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

  If Len(catTabsCsv) = 0 Or Len(outMapCsv) = 0 Then LoadTemplateSectionConfig = False: Exit Function

  catTabs  = Split(catTabsCsv, ",")
  outItems = Split(outMapCsv, ",")

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
    TrioCmd "tabfield:set_custom_property A " & Quote("SMARTSTAT=PLAYER")
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
      TrioCmd "tabfield:set_custom_property A " & Quote("SMARTSTAT=" & entityType)
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

Sub ProcessQualifierInfo(qualTab, qAliasNorm, qNorm, learn, LEARN_INI, ByRef qPrefix, ByRef qRemFrag)
  Dim qualTxt: qualTxt = ""
  If LCase(Trim(qualTab)) <> "none" And Len(Trim(qualTab)) > 0 Then
    qualTxt = CleanAfterColon(TrioCmd("page:get_property " & qualTab))
    If InStr(1, qualTxt, "ERROR:", vbTextCompare) > 0 Then qualTxt = ""
  End If

  Dim parsed, qRemainder
  parsed = NormalizeQualifierPrefixAndKey(qualTxt)
  qPrefix = "season": qRemainder = ""
  If IsArray(parsed) Then
    If Len(parsed(0)) > 0 Then qPrefix = Left(parsed(0), Len(parsed(0)) - 1)
    qRemainder = CStr(parsed(1))
  End If

  qRemFrag = ""
  If Len(Trim(qRemainder)) > 0 Then
    Dim qLeft
    If ResolveQualifierChain(qRemainder, qAliasNorm, qNorm, learn, qRemFrag, qLeft) = False Then
      If Len(Trim(qLeft)) > 0 Then
        AppendIniSectionLine LEARN_INI, "[PENDING_QUALIFIER]", Replace(qLeft, " ", "_") & " = (suggest) ??? ; QUAL remainder"
      Else
        AppendIniSectionLine LEARN_INI, "[PENDING_QUALIFIER]", NormalizeKey(qRemainder) & " = (suggest) ??? ; QUAL remainder"
      End If
      qRemFrag = ""
    End If
  End If
End Sub

Function BuildOutputTargets(outItems)
  Dim outTargets: Set outTargets = CreateObject("Scripting.Dictionary")
  Dim i, itm, parts, tgtTab, colIdx, rowIdx
  For i = LBound(outItems) To UBound(outItems)
    itm = Trim(CStr(outItems(i)))
    parts = Split(itm, ":")
    If UBound(parts) >= 2 Then
      tgtTab = Trim(parts(0))
      colIdx = CInt(parts(1))
      rowIdx = CInt(parts(2))
      If Not outTargets.Exists(colIdx) Then outTargets.Add colIdx, CreateObject("Scripting.Dictionary")
      If Not outTargets(colIdx).Exists(rowIdx) Then outTargets(colIdx).Add rowIdx, Array()
      outTargets(colIdx)(rowIdx) = PushArray(outTargets(colIdx)(rowIdx), tgtTab)
    End If
  Next
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
                  TrioCmd "tabfield:set_custom_property " & CStr(targets(j)) & " " & Quote(finalSyntax)
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
                TrioCmd "tabfield:set_custom_property " & CStr(tclear(jc)) & " " & Quote("")
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
  If Len(tmplName) = 0 Then Exit Sub

  Diag_Mark_Classify "Template detected: " & tmplName

  Dim tsec, qualTab, catTabs, outItems, filterTabs, haveFilters
  haveFilters = False
  If Not LoadTemplateSectionConfig(srcDir, tmplName, tsec, qualTab, catTabs, outItems, filterTabs, haveFilters) Then
    Diag_HardFail PHASE_03_CLASSIFY_FIELDS, "TPLCFG.MISS", "Template block missing for " & tmplName, "Add [TEMPLATE:" & tmplName & "] to SmartStat_TemplateConfig.ini"
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

  Dim qPrefix, qRemFrag
  ProcessQualifierInfo qualTab, qAliasNorm, qNorm, learn, LEARN_INI, qPrefix, qRemFrag
  Diag_Mark_DetectFilters "Qualifier processed; qPrefix=" & qPrefix

  Dim outTargets: Set outTargets = BuildOutputTargets(outItems)
  Diag_Mark_BuildOutMap "Output targets compiled"

  Dim rowCount: rowCount = DetermineRowCount(tsec, haveFilters, filterTabs)
  Dim filterFrags: filterFrags = ResolveFilterFragments(rowCount, haveFilters, filterTabs, qAliasNorm, qNorm, learn, LEARN_INI)
  Diag_Mark_DetectFilters "RowCount=" & rowCount & ", filters resolved"

  Diag_Mark_BuildSyntax "Building syntax for category columns"
  ProcessCategoryColumns catTabs, transforms, rxTransforms, learn, LEARN_INI, catAliasRaw, catRaw, catPAliasRaw, catPRaw, catAliasNorm, catPAliasNorm, catNorm, catPNorm, outTargets, rowCount, filterFrags, qPrefix, qRemFrag, entityCtx, entityType, playerSubtype
  Diag_Mark_PushToTrio "Category syntax applied to Trio custom properties"

  Diag_Mark_ApplyOverrides "Applying static overrides"
  ApplyStaticOverridesByTemplate srcDir & "SmartStat_StaticOverrides.ini", tmplName, UCase(entityCtx), UCase(playerSubtype)
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
  Dim ini: Set ini = LoadIni(staticIniPath)
  If ini Is Nothing Then Exit Sub

  Dim secSpecific, secGeneric
  secSpecific = "STATIC_FIELD_TO_SYNTAX_" & UCase(entityCtx) & "_" & UCase(tmplName)
  secGeneric  = "STATIC_FIELD_TO_SYNTAX_" & UCase(entityCtx)

  If ini.Exists(secSpecific) Then
    Dim sec: Set sec = ini(secSpecific)
    Dim k, val
    For Each k In sec.Keys
      val = CStr(sec(k))
      If UCase(entityCtx) = "PLAYER" And UCase(playerSubtype) = "P" Then
        If LCase(val) = "{{info.player.primary_position}}" Then val = "{{info.player.pitcher_hand}}"
      End If
      TrioCmd "tabfield:set_custom_property " & CStr(k) & " " & Quote(val)
    Next
  End If

  If ini.Exists(secGeneric) Then
    Dim sec2: Set sec2 = ini(secGeneric)
    Dim k2, val2
    For Each k2 In sec2.Keys
      val2 = CStr(sec2(k2))
      If UCase(entityCtx) = "PLAYER" And UCase(playerSubtype) = "P" Then
        If LCase(val2) = "{{info.player.primary_position}}" Then val2 = "{{info.player.pitcher_hand}}"
      End If
      TrioCmd "tabfield:set_custom_property " & CStr(k2) & " " & Quote(val2)
    Next
  End If

  On Error GoTo 0
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
    
    If TrioCmd("sock:socket_is_connected") Then
      TrioCmd "sock:send_socket_data on_air_get message_number=" & page_name & _
              " query=" & on_air_tabs & " message_context=" & page_desc & vbCrLf
    End If
  End If
End Function
