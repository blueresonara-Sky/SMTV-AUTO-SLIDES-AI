@echo off
setlocal EnableExtensions EnableDelayedExpansion

title SMTV Auto Slides Installer

set "SCRIPT_DIR=%~dp0"
set "SOURCE_DIR=%SCRIPT_DIR%SMTV_Slides"
set "EXTENSION_NAME=SMTV_Slides"
set "USER_DEST_ROOT=%APPDATA%\Adobe\CEP\extensions"
set "DEST_DIR=%USER_DEST_ROOT%\%EXTENSION_NAME%"
set "OLD_DEST_1=%CommonProgramFiles%\Adobe\CEP\extensions\%EXTENSION_NAME%"
set "OLD_DEST_2=%CommonProgramFiles(x86)%\Adobe\CEP\extensions\%EXTENSION_NAME%"

if not exist "%SOURCE_DIR%\CSXS\manifest.xml" (
  echo.
  echo [ERROR] Could not find the extension files in:
  echo         %SOURCE_DIR%
  echo.
  echo Make sure this BAT file is beside the SMTV_Slides folder.
  echo.
  pause
  exit /b 1
)

echo.
echo ==================================================
echo   SMTV Auto Slides Installer
echo ==================================================
echo.
echo Source:      %SOURCE_DIR%
echo Install to:  %DEST_DIR%
echo.

net session >nul 2>&1
if errorlevel 1 (
  if exist "%OLD_DEST_1%" (
    echo Requesting administrator permission to remove old Program Files install...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
  )
  if exist "!OLD_DEST_2!" (
    echo Requesting administrator permission to remove old Program Files install...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
  )
)

echo Checking Adobe CEP debug mode...
for %%K in (
  4 5 6 6.1 7 8 9 9.4
  10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28
) do (
  call :ensure_debug_mode "%%K"
)

tasklist /FI "IMAGENAME eq Adobe Premiere Pro.exe" 2>nul | find /I "Adobe Premiere Pro.exe" >nul
if not errorlevel 1 (
  echo.
  echo [WARNING] Adobe Premiere Pro appears to be running.
  echo Close Premiere before opening the updated panel if anything looks stale.
  echo.
)

if exist "%OLD_DEST_1%" (
  echo Removing old Program Files extension...
  rmdir /S /Q "%OLD_DEST_1%" >nul 2>&1
  if exist "%OLD_DEST_1%" (
    echo.
    echo [ERROR] Could not remove old extension from:
    echo         %OLD_DEST_1%
    echo.
    pause
    exit /b 1
  )
)

if exist "!OLD_DEST_2!" (
  echo Removing old Program Files ^(x86^) extension...
  rmdir /S /Q "!OLD_DEST_2!" >nul 2>&1
  if exist "!OLD_DEST_2!" (
    echo.
    echo [ERROR] Could not remove old extension from:
    echo         !OLD_DEST_2!
    echo.
    pause
    exit /b 1
  )
)

if not exist "%DEST_DIR%" (
  mkdir "%DEST_DIR%" >nul 2>&1
)

echo Installing extension files...
robocopy "%SOURCE_DIR%" "%DEST_DIR%" /MIR /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >nul
set "ROBOCOPY_EXIT=%ERRORLEVEL%"
if %ROBOCOPY_EXIT% GEQ 8 (
  echo.
  echo [ERROR] File copy failed. Robocopy exit code: %ROBOCOPY_EXIT%
  echo.
  pause
  exit /b %ROBOCOPY_EXIT%
)

echo.
echo Installation complete.
echo.
echo Next steps:
echo 1. Restart Premiere Pro.
echo 2. Open Window ^> Extensions ^> SMTV Auto Slides.
echo.
echo Press any key to close this installer window.
echo.
pause >nul
exit /b 0

:ensure_debug_mode
set "CSXS_VER=%~1"
set "REG_KEY=HKCU\Software\Adobe\CSXS.%CSXS_VER%"
set "CURRENT_VALUE="

for /f "tokens=2,*" %%A in ('reg query "%REG_KEY%" /v PlayerDebugMode 2^>nul ^| find /I "PlayerDebugMode"') do (
  set "CURRENT_VALUE=%%B"
)

if /I "!CURRENT_VALUE!"=="1" (
  echo   CSXS.!CSXS_VER!: already enabled
  goto :eof
)

reg add "%REG_KEY%" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul
if errorlevel 1 (
  echo   CSXS.!CSXS_VER!: failed to enable
) else (
  if defined CURRENT_VALUE (
    echo   CSXS.!CSXS_VER!: enabled
  ) else (
    echo   CSXS.!CSXS_VER!: created and enabled
  )
)
goto :eof
