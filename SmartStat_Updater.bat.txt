@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: ===============================================================
:: SmartStat Repo Setup/Update (Git or auto-PortableGit, no admin)
:: Repo: https://github.com/gernblanston21/SmartStat  Branch: main
:: Flags (optional):
::   set FORCE_PORTABLEGIT=1   -> skip system Git, force PortableGit
::   set DRY_RUN=1             -> resolve/tools only, no clone/pull
:: ===============================================================

:: ---- EDITABLE IF NEEDED --------------------------------------
set "REPO_URL=https://github.com/gernblanston21/SmartStat.git"
set "DEFAULT_BRANCH=main"
:: --------------------------------------------------------------

:: Flags default
if not defined FORCE_PORTABLEGIT set "FORCE_PORTABLEGIT=0"
if not defined DRY_RUN set "DRY_RUN=0"

:: Paths
set "EDRIVE=E:\"
set "ROOT=E:\EDRIVE\UNIVERSAL"
set "DEST=%ROOT%\SmartStat"
set "TOOLS=%DEST%\.tools"
set "PGIT=%TOOLS%\PortableGit"
set "PGIT_BIN=%PGIT%\bin"
set "TEMP_DL=%TEMP%\SmartStat_Installer"
if not exist "%TEMP_DL%" mkdir "%TEMP_DL%" >nul 2>&1

:: Timestamp YYYYMMDD_HHMM
for /f "tokens=1-5 delims=/:. " %%a in ("%date% %time%") do (set YYYY=%%c&set MM=%%a&set DD=%%b&set HH=%%d&set MI=%%e)
set "TS=%YYYY%%MM%%DD%_%HH%%MI%"

echo.
echo === SmartStat Installer / Updater ==========================
echo Target folder: %DEST%
echo Repository:    %REPO_URL%  (branch: %DEFAULT_BRANCH%)
echo ===========================================================
echo.

:: Require E:
if not exist "%EDRIVE%" (
  echo [ERROR] E:\ drive not found. This expects E:\EDRIVE\UNIVERSAL\SmartStat
  pause
  exit /b 1
)

:: Ensure base folders
if not exist "%ROOT%"  mkdir "%ROOT%"  >nul 2>&1
if not exist "%DEST%"  mkdir "%DEST%"  >nul 2>&1
if not exist "%TOOLS%" mkdir "%TOOLS%" >nul 2>&1

:: ---- Resolve git.exe -----------------------------------------
set "GITEXE="

if "%FORCE_PORTABLEGIT%"=="1" (
  echo [TEST] FORCE_PORTABLEGIT=1 - skipping system Git.
) else (
  for %%G in (git.exe) do set "GITEXE=%%~$PATH:G"
)

if not defined GITEXE if exist "%PGIT_BIN%\git.exe" set "GITEXE=%PGIT_BIN%\git.exe"

if not defined GITEXE (
  echo [INFO] Git not found or forced. Bootstrapping PortableGit...
  call :BootstrapPortableGit
  if errorlevel 1 (
    echo [ERROR] PortableGit bootstrap failed.
    pause
    exit /b 1
  )
  set "GITEXE=%PGIT_BIN%\git.exe"
)

echo [OK] Using git: "%GITEXE%"
"%GITEXE%" --version || (echo [ERROR] git failed to run.& pause & exit /b 1)

if "%DRY_RUN%"=="1" (
  echo [DRY-RUN] Exiting before clone/pull per DRY_RUN=1.
  goto :finish
)

:: ---- Clone or update -----------------------------------------
if exist "%DEST%\.git" (
  echo.
  echo [INFO] Existing Git repo detected. Updating...
  pushd "%DEST%" >nul
    "%GITEXE%" remote set-url origin "%REPO_URL%" 1>nul 2>nul
    "%GITEXE%" fetch --all --prune || goto :git_fail
    "%GITEXE%" checkout "%DEFAULT_BRANCH%" || goto :git_fail
    "%GITEXE%" pull --ff-only origin "%DEFAULT_BRANCH%" || goto :git_fail
  popd >nul
  echo [DONE] SmartStat updated to latest on %DEFAULT_BRANCH%.
  goto :finish
)

:: If folder has files but not a repo, back it up then clone
for /f %%A in ('dir /b "%DEST%" 2^>nul') do (
  if not exist "%DEST%\.git" (
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
)

:do_clone
echo.
echo [INFO] Cloning repository (shallow)...
"%GITEXE%" clone --depth 1 --branch "%DEFAULT_BRANCH%" "%REPO_URL%" "%DEST%"
if errorlevel 1 goto :git_fail
echo [DONE] SmartStat cloned.
goto :finish

:git_fail
echo.
echo [ERROR] Git operation failed. Check internet/firewall or branch name.
pause
exit /b 1

:finish
echo.
echo ===========================================================
echo SmartStat is ready at: %DEST%
echo Double-click this same .bat later to pull updates.
echo Normalizing permissions so operators can write logs/overrides...
echo (If this step fails silently, the script will still work for reads.)
echo ===========================================================
icacls "%DEST%" /grant Users:(OI)(CI)M /T >nul 2>&1
attrib -r "%DEST%\*.*" /s >nul 2>&1
echo Press any key to exit, or this window will close automatically in 60 seconds...
timeout /t 60 >nul
exit /b 0

:: ===============================================================
:: BootstrapPortableGit
::  - Writes a tiny PowerShell script to resolve latest asset URL
::    and download/extract PortableGit to %PGIT%
:: ===============================================================
:BootstrapPortableGit
setlocal
set "OUT=%TEMP_DL%\PortableGit.7z.exe"
set "PS1=%TEMP_DL%\bootstrap_pgit.ps1"
set "URLTXT=%TEMP_DL%\pgit_url.txt"
del /f /q "%OUT%" "%PS1%" "%URLTXT%" >nul 2>&1

:: Detect arch (64-bit unless truly 32-bit)
set "ARCH=64-bit"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set "ARCH=32-bit"

:: --- Write the PowerShell helper script line-by-line (no () block) ---
>>"%PS1%" echo $ErrorActionPreference='Stop'
>>"%PS1%" echo $ProgressPreference='SilentlyContinue'
>>"%PS1%" echo $arch='%ARCH%'
>>"%PS1%" echo $headers=@{ 'User-Agent'='SmartStat-Installer' }
>>"%PS1%" echo $api='https://api.github.com/repos/git-for-windows/git/releases/latest'
>>"%PS1%" echo try { $rel=Invoke-RestMethod -Headers $headers -Uri $api } catch { $rel=$null }
>>"%PS1%" echo if ($arch -eq '64-bit') { $pat='PortableGit-.*-64-bit\.7z\.exe$' } else { $pat='PortableGit-.*-32-bit\.7z\.exe$' }
>>"%PS1%" echo $url=$null
>>"%PS1%" echo if ($rel) { $url = ($rel.assets ^| Where-Object { $_.name -match $pat } ^| Select-Object -First 1 -ExpandProperty browser_download_url) }
>>"%PS1%" echo if (-not $url) { if ($arch -eq '64-bit') { $url='https://github.com/git-for-windows/git/releases/download/v2.51.2.windows.1/PortableGit-2.51.2-64-bit.7z.exe' } else { $url='https://github.com/git-for-windows/git/releases/download/v2.51.2.windows.1/PortableGit-2.51.2-32-bit.7z.exe' } }
>>"%PS1%" echo Set-Content -Path '%URLTXT%' -Value $url -NoNewline
>>"%PS1%" echo Invoke-WebRequest -Headers $headers -UseBasicParsing -Uri $url -OutFile '%OUT%'
>>"%PS1%" echo if (-not (Test-Path '%PGIT%')) { New-Item -ItemType Directory -Path '%PGIT%' ^| Out-Null }
>>"%PS1%" echo Start-Process -FilePath '%OUT%' -ArgumentList '-y','-o%PGIT%' -Wait
>>"%PS1%" echo if (-not (Test-Path '%PGIT%\bin\git.exe')) { exit 2 } else { exit 0 }

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
set "RC=%ERRORLEVEL%"

if exist "%URLTXT%" (
  set /p PGIT_URL=<"%URLTXT%"
  if defined PGIT_URL echo [INFO] PortableGit URL: %PGIT_URL%
)

if not "%RC%"=="0" (
  echo [ERROR] PowerShell bootstrap returned %RC%.
  endlocal & exit /b 1
)

if exist "%PGIT_BIN%\git.exe" (
  echo [OK] PortableGit ready at: %PGIT_BIN%
  endlocal & exit /b 0
)

echo [ERROR] PortableGit extraction missing git.exe
endlocal & exit /b 1
