@echo off
setlocal EnableExtensions EnableDelayedExpansion

if /I "%~1"=="noelevate" goto :skip_elevation

:: ==========================================================
:: UAC / Admin Elevation
::  - If not admin, relaunch this script elevated.
::  - The original window will exit; work happens in new window.
:: ==========================================================
whoami /groups | find "S-1-5-32-544" >nul 2>&1
if errorlevel 1 (
  echo [INFO] Not running as admin. Attempting elevation...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList 'elevated' -Verb RunAs"
  exit /b
)

:: If we were launched with "elevated", drop that arg
if /I "%~1"=="elevated" shift

echo [OK] Running with admin privileges.
echo.

:skip_elevation

:: ====== Editable settings ======
set "REPO_URL=https://github.com/gernblanston21/SmartStat.git"
set "DEFAULT_BRANCH=main"
if not defined FORCE_PORTABLEGIT set "FORCE_PORTABLEGIT=0"
if not defined DRY_RUN set "DRY_RUN=0"

:: ====== Paths ======
set "EDRIVE=E:\"
set "ROOT=%EDRIVE%EDRIVE\UNIVERSAL"
set "DEST=%ROOT%\SmartStat"
set "TOOLS=%ROOT%\.tools"
set "PGIT=%TOOLS%\PortableGit"
set "PGIT_BIN=%PGIT%\bin"
set "TEMP_DL=%TEMP%\SmartStat_Installer"
if not exist "%TEMP_DL%" mkdir "%TEMP_DL%" >nul 2>&1

:: ====== Timestamp ======
for /f "tokens=1-5 delims=/:. " %%a in ("%date% %time%") do (
  set YYYY=%%c
  set MM=%%a
  set DD=%%b
  set HH=%%d
  set MI=%%e
)
set "TS=%YYYY%%MM%%DD%_%HH%%MI%"

:: make sure base exists early so we can log
if not exist "%DEST%" mkdir "%DEST%" >nul 2>&1

set "LOG=%ROOT%\SmartStat_Installer_Log_%TS%.txt"

echo ============================================================================
echo SmartStat Updater Diagnostic Log - %TS%
echo Target: %DEST%
echo Repo:   %REPO_URL% (branch: %DEFAULT_BRANCH%)
echo ============================================================================

>>"%LOG%" echo ============================================================================
>>"%LOG%" echo SmartStat Updater Diagnostic Log - %TS%
>>"%LOG%" echo Target: %DEST%
>>"%LOG%" echo Repo:   %REPO_URL% (branch: %DEFAULT_BRANCH%)
>>"%LOG%" echo ============================================================================

echo.
echo Running diagnostic updater. Detailed log: %LOG%
echo.

:: ====== Env snapshot ======
>>"%LOG%" echo --- Environment snapshot ---
>>"%LOG%" echo USER: %USERNAME%
>>"%LOG%" echo COMPUTER: %COMPUTERNAME%
>>"%LOG%" echo OS: %OS%
>>"%LOG%" echo PROCESSOR_ARCHITECTURE: %PROCESSOR_ARCHITECTURE%
>>"%LOG%" echo TEMP: %TEMP%
>>"%LOG%" echo PATH: %PATH%
>>"%LOG%" echo HTTP_PROXY: %HTTP_PROXY%
>>"%LOG%" echo HTTPS_PROXY: %HTTPS_PROXY%
>>"%LOG%" echo.

:: ====== Check E: ======
if not exist "%EDRIVE%" (
  echo [ERROR] E:\ drive not found. This expects E:\EDRIVE\UNIVERSAL\SmartStat
  >>"%LOG%" echo [ERROR] E:\ drive not found.
  pause
  exit /b 1
)

if not exist "%ROOT%"  mkdir "%ROOT%"  >nul 2>&1
if not exist "%DEST%"  mkdir "%DEST%"  >nul 2>&1
if not exist "%TOOLS%" mkdir "%TOOLS%" >nul 2>&1

:: ====== Resolve git ======
set "GITEXE="

if "%FORCE_PORTABLEGIT%"=="1" (
  echo [TEST] FORCE_PORTABLEGIT=1 - skipping system Git.
  >>"%LOG%" echo [TEST] FORCE_PORTABLEGIT=1 - skipping system Git.
) else (
  :: try to find git in PATH with extra guard
  for %%G in (git.exe) do (
    if "%%~$PATH:G" NEQ "" (
      set "GITEXE=%%~$PATH:G"
    )
  )
  if defined GITEXE (
    echo [INFO] Found system git: "%GITEXE%"
    >>"%LOG%" echo [INFO] Found system git: "%GITEXE%"
  ) else (
    echo [INFO] system git not found in PATH.
    >>"%LOG%" echo [INFO] system git not found in PATH.
  )
)

:: if we skipped system git, or none found, try existing portable first
if not defined GITEXE if exist "%PGIT_BIN%\git.exe" (
  set "GITEXE=%PGIT_BIN%\git.exe"
  echo [INFO] Using existing PortableGit: "%GITEXE%"
  >>"%LOG%" echo [INFO] Using existing PortableGit: "%GITEXE%"
)

:: if still no git, bootstrap it
if not defined GITEXE (
  echo [INFO] Git not found. Attempting PortableGit bootstrap (PowerShell)
  >>"%LOG%" echo [INFO] Git not found. Attempting PortableGit bootstrap (PowerShell)
  call :BootstrapPortableGit "%LOG%"
  set "RC=%ERRORLEVEL%"
  >>"%LOG%" echo [INFO] PortableGit bootstrap returned %RC%.
  if not "%RC%"=="0" (
    echo [ERROR] PortableGit bootstrap failed. See log: %LOG%
    pause
    exit /b 1
  )
  if exist "%PGIT_BIN%\git.exe" set "GITEXE=%PGIT_BIN%\git.exe"
)

if not defined GITEXE (
  echo [ERROR] No git available after bootstrap. See log: %LOG%
  >>"%LOG%" echo [ERROR] No git available after bootstrap.
  pause
  exit /b 1
)

echo [OK] Using git: "%GITEXE%"
>>"%LOG%" echo [OK] Using git: "%GITEXE%"

:: ====== Check git works ======
"%GITEXE%" --version >>"%LOG%" 2>&1 || (
  echo [ERROR] git --version failed. See log: %LOG%
  >>"%LOG%" echo [ERROR] git --version failed.
  pause
  exit /b 1
)

if "%DRY_RUN%"=="1" (
  echo [DRY-RUN] Exiting before clone/pull per DRY_RUN=1.
  >>"%LOG%" echo [DRY-RUN] Exiting before clone/pull per DRY_RUN=1.
  goto :finish
)

:: ====== Clone or update ======
if exist "%DEST%\.git" (
  echo [INFO] Existing Git repo detected. Updating...
  >>"%LOG%" echo [INFO] Existing Git repo detected. Updating...
  pushd "%DEST%" >nul
    "%GITEXE%" remote set-url origin "%REPO_URL%" >>"%LOG%" 2>&1
    "%GITEXE%" fetch --all --prune --progress >>"%LOG%" 2>&1 || goto :git_fail
    "%GITEXE%" checkout "%DEFAULT_BRANCH%" >>"%LOG%" 2>&1 || goto :git_fail
    "%GITEXE%" pull --ff-only origin "%DEFAULT_BRANCH%" --progress >>"%LOG%" 2>&1 || goto :git_fail
  popd >nul
  echo [DONE] SmartStat updated to latest on %DEFAULT_BRANCH%.
  >>"%LOG%" echo [DONE] SmartStat updated to latest on %DEFAULT_BRANCH%.
  goto :finish
)

:: If folder has real files (not just .tools) but not a repo, back it up then clone
set "HAS_REAL_CONTENT="

for /f "delims=" %%A in ('dir /b "%DEST%" 2^>nul') do (
  rem Ignore our own tools folder when deciding if this is a "non-git copy"
  if /I not "%%A"==".tools" (
    set "HAS_REAL_CONTENT=1"
  )
)

if defined HAS_REAL_CONTENT if not exist "%DEST%\.git" (
  set "BKP=%ROOT%\SmartStat_backup_%TS%"
  echo.
  echo [INFO] Converting existing non-git copy to real clone.
  echo [INFO] Backing up to: "%BKP%"
  mkdir "%BKP%" >nul 2>&1
  robocopy "%DEST%" "%BKP%" /E /R:1 /W:1 >nul
  rmdir /S /Q "%DEST%" >nul 2>&1
  mkdir "%DEST%" >nul 2>&1
  goto :do_clone
)

echo [DEBUG] Checking paths before clone...>>"%LOG%"
echo [DEBUG] DEST="%DEST%">>"%LOG%"
dir "%DEST%" >>"%LOG%" 2>&1
dir "%ROOT%" >>"%LOG%" 2>&1
dir "E:\EDRIVE" >>"%LOG%" 2>&1

:do_clone
echo [INFO] Cloning repository (shallow)...
>>"%LOG%" echo [INFO] Cloning repository (shallow)...
set "ATTEMPTS=0"
:clone_try
set /a ATTEMPTS+=1
"%GITEXE%" clone --depth 1 --branch "%DEFAULT_BRANCH%" "%REPO_URL%" "%DEST%" >>"%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
if "%RC%"=="0" (
  echo [DONE] SmartStat cloned.
  >>"%LOG%" echo [DONE] SmartStat cloned.
  goto :finish
)
echo [WARN] Clone attempt %ATTEMPTS% failed with exit code %RC%.
>>"%LOG%" echo [WARN] Clone attempt %ATTEMPTS% failed with exit code %RC%.
if %ATTEMPTS% lss 3 (
  timeout /t 5 >nul
  goto :clone_try
)
goto :git_fail

:git_fail
echo [ERROR] Git operation failed. Detailed log: %LOG%
>>"%LOG%" echo [ERROR] Git operation failed.
type "%LOG%" | more
pause
exit /b 1

:finish
echo.
echo ===========================================================
echo SmartStat is ready at: %DEST%
echo Diagnostic log: %LOG%
echo ===========================================================
icacls "%DEST%" /grant Users:(OI)(CI)M /T >nul 2>&1
attrib -r "%DEST%\*.*" /s >nul 2>&1
echo Press any key to exit...
pause
exit /b 0

:: ===========================================================
:: BootstrapPortableGit  (no parentheses!)
:: arg %1 = log file path
:: ===========================================================
:BootstrapPortableGit
setlocal
set "LOGFILE=%~1"
if not defined LOGFILE set "LOGFILE=%TEMP%\SmartStat_Bootstrap_Log.txt"

set "OUT=%TEMP_DL%\PortableGit.7z.exe"
set "PS1=%TEMP_DL%\bootstrap_pgit.ps1"
set "URLTXT=%TEMP_DL%\pgit_url.txt"
del /f /q "%OUT%" "%PS1%" "%URLTXT%" >nul 2>&1

:: Detect arch
set "ARCH=64-bit"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set "ARCH=32-bit"

>>"%LOGFILE%" echo [BOOT] Writing PowerShell helper: %PS1%

:: line-by-line to avoid parser errors
>>"%PS1%" echo $ErrorActionPreference='Stop'
>>"%PS1%" echo $ProgressPreference='SilentlyContinue'
>>"%PS1%" echo $arch='%ARCH%'
>>"%PS1%" echo $headers=@{ 'User-Agent'='SmartStat-Installer' }
>>"%PS1%" echo $api='https://api.github.com/repos/git-for-windows/git/releases/latest'
>>"%PS1%" echo try { $rel=Invoke-RestMethod -Headers $headers -Uri $api } catch { $rel=$null; Write-Output 'WARN: Release lookup failed.' }
>>"%PS1%" echo if ($arch -eq '64-bit') { $pat='PortableGit-.*-64-bit\.7z\.exe$' } else { $pat='PortableGit-.*-32-bit\.7z\.exe$' }
>>"%PS1%" echo $url=$null
>>"%PS1%" echo if ($rel) { $asset = $rel.assets ^| Where-Object { $_.name -match $pat } ^| Select-Object -First 1 }
>>"%PS1%" echo if ($asset) { $url = $asset.browser_download_url }
>>"%PS1%" echo if (-not $url) {
>>"%PS1%" echo   Write-Output 'INFO: Falling back to pinned PortableGit URL.'
>>"%PS1%" echo   if ($arch -eq '64-bit') { $url='https://github.com/git-for-windows/git/releases/download/v2.51.2.windows.1/PortableGit-2.51.2-64-bit.7z.exe' } else { $url='https://github.com/git-for-windows/git/releases/download/v2.51.2.windows.1/PortableGit-2.51.2-32-bit.7z.exe' }
>>"%PS1%" echo }
>>"%PS1%" echo Set-Content -Path '%URLTXT%' -Value $url -NoNewline
>>"%PS1%" echo Write-Output "Downloading PortableGit from: $url"

:: Prefer curl.exe if available, otherwise Invoke-WebRequest
>>"%PS1%" echo $curl = "$env:SystemRoot\System32\curl.exe"
>>"%PS1%" echo if (Test-Path $curl) {
>>"%PS1%" echo   ^& $curl -L -o '%OUT%' $url
>>"%PS1%" echo } else {
>>"%PS1%" echo   Invoke-WebRequest -Headers $headers -Uri $url -OutFile '%OUT%'
>>"%PS1%" echo }

>>"%PS1%" echo if (-not (Test-Path '%PGIT%')) {
>>"%PS1%" echo   New-Item -ItemType Directory -Path '%PGIT%' ^| Out-Null
>>"%PS1%" echo }
>>"%PS1%" echo Start-Process -FilePath '%OUT%' -ArgumentList '-y','-o%PGIT%' -Wait
>>"%PS1%" echo if (-not (Test-Path '%PGIT%\bin\git.exe')) { exit 2 } else { exit 0 }

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
set "RC=%ERRORLEVEL%"

if exist "%URLTXT%" (
  set /p PGIT_URL=<"%URLTXT%"
  if defined PGIT_URL (
    echo [INFO] PortableGit URL: %PGIT_URL%
    >>"%LOGFILE%" echo [INFO] PortableGit URL: %PGIT_URL%
  )
)

if not "%RC%"=="0" (
  echo [ERROR] PowerShell bootstrap returned %RC%.
  >>"%LOGFILE%" echo [ERROR] PowerShell bootstrap returned %RC%.
  endlocal & exit /b %RC%
)

if exist "%PGIT_BIN%\git.exe" (
  echo [OK] PortableGit ready at: %PGIT_BIN%
  >>"%LOGFILE%" echo [OK] PortableGit ready at: %PGIT_BIN%
  endlocal & exit /b 0
)

echo [ERROR] PortableGit extraction missing git.exe
>>"%LOGFILE%" echo [ERROR] PortableGit extraction missing git.exe
endlocal & exit /b 1