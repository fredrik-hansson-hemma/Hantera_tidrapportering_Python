@echo off
title Chrome – Antura-session

:: ============================================================
:: Profilen sparas i mappen "chrome_profile" i projektmappen.
:: Första gången: logga in på Antura. Sessionen sparas sedan.
:: ============================================================

set PORT=9222
set PROFILE=%~dp0chrome_profile
set URL=https://ucrantura.ucr.uu.se/Views/MyPlace/TimeReporting/Reporting.aspx

:: Sök efter Chrome på vanliga ställen
set CHROME=
for %%P in (
    "%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"
    "%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"
    "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%P (
        if not defined CHROME set CHROME=%%P
    )
)

if not defined CHROME (
    echo.
    echo *** Chrome hittades inte pa nagot standardstalle ***
    echo Redigera start_chrome.bat och ange den ratta sokvägen till chrome.exe
    echo.
    pause
    exit /b 1
)

echo.
echo Startar Chrome pa port %PORT% ...
echo Profil: %PROFILE%
echo.
echo Logga in pa Antura om det behovs, kör sedan:
echo   poetry run python antura_filler.py
echo.

%CHROME% ^
  --remote-debugging-port=%PORT% ^
  --user-data-dir="%PROFILE%" ^
  "%URL%"
